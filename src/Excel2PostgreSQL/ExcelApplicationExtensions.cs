using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2PostgreSQL
{
    internal static class ExcelApplicationExtensions
    {
        public static Excel.Workbook OpenWorkbook(this Excel.Application application, string filename)
        {
            var workbooks = application.Workbooks;
            try
            {
                var workbook = workbooks.Open(filename, ReadOnly: true);
                return workbook;
            }
            finally
            {
                Marshal.ReleaseComObject(workbooks);
            }
        }
    }
}
