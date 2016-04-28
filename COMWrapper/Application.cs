using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace COMWrapper
{
    public class Application : IDisposable
    {
        Excel.Application _app;
        List<Workbook> _wbs;

        public Application()
        {
            _app = new Excel.Application();
            _wbs = new List<Workbook>();
        }

        // All of the following private enums are poorly documented
        private enum XlCorruptLoad
        {
            NormalLoad = 0,
            RepairFile = 1,
            ExtractData = 2
        }

        private enum XlUpdateLinks
        {
            Yes = 2,
            No = 0
        }

        private enum XlPlatform
        {
            Macintosh = 1,
            Windows = 2,
            MSDOS = 3
        }

        public Workbook OpenWorkbook(string relpath)
        {
            // get the absolute path
            var abspath = System.IO.Path.GetFullPath(relpath);

            // we need to disable all alerts, e.g., password prompts, etc.
            _app.DisplayAlerts = false;

            // disable macros
            _app.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            // This call is stupid.  See:
            // http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.workbooks.open%28v=office.11%29.aspx
            _app.Workbooks.Open(abspath,                    // FileName (String)
                               XlUpdateLinks.Yes,           // UpdateLinks (XlUpdateLinks enum)
                               true,                        // ReadOnly (Boolean)
                               Missing.Value,               // Format (int?)
                               "thisisnotapassword",        // Password (String)
                               Missing.Value,               // WriteResPassword (String)
                               true,                        // IgnoreReadOnlyRecommended (Boolean)
                               Missing.Value,               // Origin (XlPlatform enum)
                               Missing.Value,               // Delimiter; if the filetype is txt (String)
                               Missing.Value,               // Editable; not what you think (Boolean)
                               false,                       // Notify (Boolean)
                               Missing.Value,               // Converter(int)
                               false,                       // AddToMru (Boolean)
                               Missing.Value,               // Local; really "use my locale?" (Boolean)
                               XlCorruptLoad.RepairFile);   // CorruptLoad (XlCorruptLoad enum)

            // init wrapped workbook
            // TODO: the array index here really should depend on the number of open workbooks
            var wb = new Workbook(_app.Workbooks[1], _app);

            // add to list
            _wbs.Add(wb);

            return wb;
        }

        public void Dispose()
        {
            foreach (var wb in _wbs)
            {
                wb.Dispose();
            }
            _app.Quit();
            Marshal.ReleaseComObject(_app);
            _app = null;
        }
    }
}
