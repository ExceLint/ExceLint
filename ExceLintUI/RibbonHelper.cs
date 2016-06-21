using Excel = Microsoft.Office.Interop.Excel;

namespace ExceLintUI
{
    static class RibbonHelper
    {
        public static Excel.Worksheet GetWorksheetByName(string name, Excel.Sheets sheets)
        {
            foreach (Excel.Worksheet ws in sheets)
            {
                if (ws.Name == name)
                {
                    return ws;
                }
            }
            return null;
        }
    }
}
