using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using NpgsqlTypes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2PostgreSQL
{
    internal class TransferTable
    {
        private readonly ColumnInfo[] _columns;
        private readonly TransferRow[] _rows;

        private TransferTable(string name, IEnumerable<ColumnInfo> columns, IEnumerable<TransferRow> rows)
        {
            Name = name;
            _columns = columns.ToArray();
            _rows = rows.ToArray();
        }

        public IEnumerable<ColumnInfo> Columns => _columns;
        public IEnumerable<TransferRow> Rows => _rows;
        public string Name { get; }

        public static TransferTable FromWorkSheet(Excel.Worksheet sheet)
        {
            List<ColumnInfo> columns = new List<ColumnInfo>();
            List<TransferRow> rows = new List<TransferRow>();

            sheet.Activate();
            Excel.Range usedRange = sheet.UsedRange;
            Excel.Range usedRows = usedRange.Rows;
            Excel.Range usedColumns = usedRange.Columns;
            for (int i = 1; i <= usedColumns.Count; i++)
            {
                Excel.Range cell;
                ColumnInfo ci = null;
                cell = usedColumns.Cells[1, i];
                ci = new ColumnInfo(i - 1, cell.Value2 as string ?? $"Column {i}");
                Marshal.ReleaseComObject(cell);
                ColumnTypeInferer ct = new ColumnTypeInferer();
                for (int j = 2; j <= usedRows.Count; j++)
                {
                    if (i == 1)
                    {
                        rows.Add(new TransferRow(new object[usedColumns.Count]));
                    }
                    cell = usedRange.Cells[j, i];
                    var value = cell.Value2;
                    if (value == null)
                    {
                        ci.Nullable = true;
                    }
                    ct.UpdateType(value, cell.NumberFormat);
                    rows[j - 2][ci] = value;
                    Marshal.ReleaseComObject(cell);
                }
                ci.Type = ct.ResultTypeMightBeDateOrTime ? NpgsqlDbType.Timestamp : ct.ResultType;
                columns.Add(ci);
            }
            var table = new TransferTable(sheet.Name, columns, rows);
            Marshal.ReleaseComObject(usedRows);
            Marshal.ReleaseComObject(usedColumns);
            Marshal.ReleaseComObject(usedRange);
            return table;
        }
    }
}
