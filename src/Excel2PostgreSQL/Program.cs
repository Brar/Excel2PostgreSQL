using Microsoft.Extensions.CommandLineUtils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Excel2PostgreSQL.Properties;

namespace Excel2PostgreSQL
{
    public class Program
    {
        internal static CommandLineApplication Args { get; } = new CommandLineApplication();

        public static int Main(string[] args)
        {
            SetupArgsParser();
            return Args.Execute(args);
        }

        private static void SetupArgsParser()
        {
            string applicationName = Path.GetFileNameWithoutExtension(Environment.GetCommandLineArgs()[0]);
            string applicationVersion = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion;
            Args.Name = applicationName;
            Args.Argument(LocalizedStrings.FileArgName, LocalizedStrings.FileArgDescription);
            Args.HelpOption("-h | -? | --help");
            Args.VersionOption("-V | --version", $"{applicationName} v{applicationVersion}");
            Args.OnExecute(() => ExecuteCommandLineApplication());
        }

        private static int ExecuteCommandLineApplication()
        {
            if (Args.Arguments.First().Value == null)
            {
                Console.Error.WriteLine(LocalizedStrings.MissingExcelFileArg);
                Console.Error.WriteLine();
                Console.Error.WriteLine(Args.GetHelpText());
                return 1;
            }
            var app = new Excel.Application { Visible = true };
            var wb = app.OpenWorkbook(Path.GetFullPath(Args.Arguments[0].Value));
            IEnumerable<Table> tableInfos = GetTableInfos(wb);
            var dbName = wb.Name.Substring(0, wb.Name.Length - 5);
            Console.WriteLine(LocalizedStrings.CreatingDatabase, dbName);

            Database.Connect();
            Database.CreateDb(dbName, true);
            Database.Connect(dbName);

            foreach (var table in tableInfos)
            {
                Console.WriteLine("\t" + LocalizedStrings.CreatingTable, table.Name);
                Database.AddTableWithData(table);
            }
            Database.Disconnect();
            wb.Close(false);
            Marshal.ReleaseComObject(wb);
            app.Quit();
            Marshal.ReleaseComObject(app);
            return 0;
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
