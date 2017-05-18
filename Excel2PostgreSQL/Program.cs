using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2PostgreSQL
{
    class Program
    {
        static void Main(string[] args)
        {
            var app = new Excel.Application {Visible = true};
            var wb = app.OpenWorkbook(Path.GetFullPath(args[0]));
            IEnumerable<Table> tableInfos = GetTableInfos(wb);
            var dbName = wb.Name.Substring(0, wb.Name.Length - 5);
            Console.WriteLine($"Erstelle Datenbank \"{dbName}\"");

            Database.Connect();
            Database.CreateDb(dbName, true);
            Database.Connect(dbName);

            foreach (var table in tableInfos)
            {
                Console.WriteLine($"\tErstelle Tabelle \"{table.Name}\"");
                Database.AddTableWithData(table);
            }
            Database.Disconnect();
            wb.Close(false);
            Marshal.ReleaseComObject(wb);
            app.Quit();
            Marshal.ReleaseComObject(app);
        }

        private static IEnumerable<Table> GetTableInfos(Excel.Workbook wb)
        {
            var sheets = wb.Worksheets;
            try
            {
                List<Table> tables = new List<Table>(sheets.Count);
                foreach (Excel.Worksheet sheet in sheets)
                {
                    tables.Add(Table.FromWorkSheet(sheet));
                    Marshal.ReleaseComObject(sheet);
                }
                return tables;
            }
            finally
            {
                Marshal.ReleaseComObject(sheets);
            }
        }
    }
}
