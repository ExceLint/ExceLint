using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Depends;

namespace COMWrapper
{
    public class Workbook : IDisposable
    {
        private Excel.Application _app;
        private Excel.Workbook _wb;

        public Workbook(Excel.Workbook wb, Excel.Application app)
        {
            _app = app;
            _wb = wb;
        }

        public void Dispose()
        {
            _wb.Close();
            Marshal.ReleaseComObject(_wb);
            _wb = null;
        }

        public DAG buildDependenceGraph()
        {
            return new DAG(_wb, _app, true);
        }
    }
}
