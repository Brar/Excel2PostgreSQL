using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.CSharp.RuntimeBinder;
using NpgsqlTypes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2PostgreSQL
{
    internal class Table
    {
        public static Table FromWorkSheet(Excel.Worksheet sheet)
        {
            sheet.Activate();
            Excel.Range usedRange = sheet.UsedRange;
            var infos = InitializeColumnInfos(usedRange);
            object[][] data = GetData(usedRange, ref infos);
            return new Table(sheet.Name, infos, data);
        }

        private static object[][] GetData(Excel.Range usedRange, ref ColumnInfo[] infos)
        {
            Excel.Range usedRows = usedRange.Rows;
            Excel.Range usedColumns = usedRange.Columns;
            object[][] data = new object[usedRows.Count - 1][];
            for (int i = 2; i <= usedRows.Count; i++)
            {
                data[i - 2] = new object[usedColumns.Count];
                for (int j = 1; j <= usedColumns.Count; j++)
                {
                    Excel.Range cell = usedRange.Cells[i, j];
                    Excel.DisplayFormat format = cell.DisplayFormat;
                    var value = cell.Value2;
                    if (value == null)
                    {
                        infos[j - 1].Nullable = true;
                    }
                    infos[j - 1].Type = InferType(value, format.NumberFormat, infos[j - 1].Type);
                    data[i - 2][j - 1] = value;
                    Marshal.ReleaseComObject(format);
                    Marshal.ReleaseComObject(cell);
                }

            }
            Marshal.ReleaseComObject(usedRows);
            Marshal.ReleaseComObject(usedColumns);
            return data;
        }

        private static NpgsqlDbType InferType(dynamic value, string numberFormat, NpgsqlDbType previousType)
        {
            if (previousType == NpgsqlDbType.Text)
            {
                return NpgsqlDbType.Text;
            }
            if (value is double)
            {
                if (previousType == NpgsqlDbType.Double)
                {
                    return NpgsqlDbType.Double;
                }
                if (previousType == NpgsqlDbType.Timestamp)
                {
                    return NpgsqlDbType.Timestamp;
                }
                double doubleValue = value;
                if (numberFormat.Contains('$'))
                {
                    return NpgsqlDbType.Money;
                }
                bool hasDatePart = numberFormat.Contains('d') ||
                                   numberFormat.Contains('m') ||
                                   numberFormat.Contains('y');
                bool hasTimePart = numberFormat.Contains('h') ||
                                   numberFormat.Contains('m') ||
                                   numberFormat.Contains('y');
                if (Math.Abs(doubleValue % 1d) < Double.Epsilon)
                {
                    if (previousType == NpgsqlDbType.Time)
                    {
                        return NpgsqlDbType.Timestamp;
                    }
                    if (hasDatePart)
                    {
                       return NpgsqlDbType.Date;
                    }
                    if (previousType == NpgsqlDbType.Bigint)
                    {
                        return NpgsqlDbType.Bigint;
                    }
                    if (previousType == NpgsqlDbType.Integer)
                    {
                        return NpgsqlDbType.Integer;
                    }
                    double intValue = doubleValue > 0d ? Math.Floor(doubleValue) : Math.Ceiling(doubleValue);
                    if (intValue < short.MaxValue || intValue > short.MinValue)
                    {
                        return NpgsqlDbType.Smallint;
                    }
                    if (intValue < int.MaxValue || intValue > int.MinValue)
                    {
                        return NpgsqlDbType.Integer;
                    }
                    if (intValue < long.MaxValue || intValue > long.MinValue)
                    {
                        return NpgsqlDbType.Bigint;
                    }
                }
                if (previousType == NpgsqlDbType.Date || hasDatePart)
                {
                    return NpgsqlDbType.Timestamp;
                }
                if (hasTimePart)
                {
                    return NpgsqlDbType.Time;
                }
                return NpgsqlDbType.Double;                
            }
            if (value is bool)
            {
                return NpgsqlDbType.Boolean;
            }
            return NpgsqlDbType.Text;
        }

        private static ColumnInfo[] InitializeColumnInfos(Excel.Range usedRange)
        {
            Excel.Range usedColumns = usedRange.Columns;
            ColumnInfo[] infos = new ColumnInfo[usedColumns.Count];
            for (int i = 1; i <= usedColumns.Count; i++)
            {
                Excel.Range cell = usedColumns.Cells[1, i];
                var value = cell.Value2 as string ?? $"Column {i}";
                infos[i - 1] = new ColumnInfo { Name = value };
                Marshal.ReleaseComObject(cell);
            }
            Marshal.ReleaseComObject(usedColumns);
            return infos;
        }

        private Table(string name, IList<ColumnInfo> columnInfos, IEnumerable<object[]> data)
        {
            Name = name;
            ColumnInfos = columnInfos;
            Data = data;
        }

        public string Name { get; set; }
        public IList<ColumnInfo> ColumnInfos { get; }
        public IEnumerable<object[]> Data { get; }
    }
}
