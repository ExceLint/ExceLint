using System.Collections.Generic;
using System.Collections.Concurrent;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExceLintUI
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override object RequestComAddInAutomationService()
        {
            if (Globals.Ribbons.ExceLintRibbon.CurrentWorkbook == null)
            {
                //return new WorkbookState(Globals.ThisAddIn.Application, Globals.ThisAddIn.Application.ActiveWorkbook);
                return null;
            } else
            {
                return Globals.Ribbons.ExceLintRibbon.CurrentWorkbook;
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
