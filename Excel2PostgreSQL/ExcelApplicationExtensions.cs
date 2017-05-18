using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
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
                var workbook = workbooks.Open(filename);
                return workbook;
            }
            finally
            {
                Marshal.ReleaseComObject(workbooks);
            }
        }
    }
}
