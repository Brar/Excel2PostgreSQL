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
        internal static CommandLineApplication Args { get; private set; }

        public static int Main(string[] args)
        {
            SetupArgsParser();
            return Args.Execute(args);
        }

        private static void SetupArgsParser()
        {
            string applicationName = Path.GetFileNameWithoutExtension(Environment.GetCommandLineArgs()[0]);
            string applicationVersion = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion;
            Args = new CommandLineApplication { Name = applicationName };
            Args.Argument(LocalizedStrings.FileArgName, LocalizedStrings.FileArgDescription);
            Args.HelpOption("-h | -? | --help");
            Args.VersionOption("-V | --version", $"{applicationName} v{applicationVersion}");
            Args.OnExecute(() => ExecuteCommandLineApplication());
        }

        private static int ExecuteCommandLineApplication()
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            try
            {
                var excelFile = Args.Arguments.First().Value;

                if (excelFile == null)
                {
                    return HandleError(ExitCode.ErrorMissingExcelFileArg, LocalizedStrings.MissingExcelFileArg);
                }

                excelFile = Path.GetFullPath(excelFile);

                if (!File.Exists(excelFile))
                {
                    return HandleError(ExitCode.ErrorFileDoesNotExist, LocalizedStrings.FileDoesNotExist, excelFile);
                }
                app = new Excel.Application { Visible = false };
                wb = app.OpenWorkbook(excelFile);
                return ImportWorkbook(wb);
            }
            catch (Exception e)
            {
                return HandleUnexpectedException(e);
            }
            finally
            {
                CleanUpWorkbook(wb);
                CleanUpApplication(app);
            }
        }

        private static int ImportWorkbook(Excel.Workbook wb)
        {
            IEnumerable<TransferTable> tableInfos = GetTableInfos(wb);
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
            return (int)ExitCode.Success;
        }

        private static int HandleUnexpectedException(Exception e)
        {
#if DEBUG
            Console.Error.WriteLine("An unexpected error occured: {0}", e.ToString());
#else
            Console.Error.WriteLine("An unexpected error occured: {0}", e.Message);
#endif
            return (int)ExitCode.ErrorUnknown;
        }

        private static int HandleError(ExitCode exitCode, string message, params object[] arg)
        {
            Console.Error.WriteLine(message, arg);
            if (exitCode == ExitCode.ErrorMissingExcelFileArg)
            {
                Console.Error.WriteLine();
                Console.Error.WriteLine(Args.GetHelpText());
            }
            return (int)exitCode;
        }

        private static void CleanUpApplication(Excel.Application app)
        {
            if (app != null)
            {
                app.Quit();
                Marshal.FinalReleaseComObject(app);
            }
        }

        private static void CleanUpWorkbook(Excel.Workbook wb)
        {
            if (wb != null)
            {
                wb.Close(false);
                Marshal.FinalReleaseComObject(wb);
            }
        }

        private static IEnumerable<TransferTable> GetTableInfos(Excel.Workbook wb)
        {
            var sheets = wb.Worksheets;
            try
            {
                List<TransferTable> tables = new List<TransferTable>(sheets.Count);
                foreach (Excel.Worksheet sheet in sheets)
                {
                    tables.Add(TransferTable.FromWorkSheet(sheet));
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
