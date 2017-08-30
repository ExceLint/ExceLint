using System;
using System.Collections.Generic;
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

        public static int ArgMax<T> (T[] array) where T : IComparable<T>
        {
            int max = 0;
            for (int i = 0; i < array.Length; i++)
            {
                // generic "greater than"
                if (array[i].CompareTo(array[max]) > 0)
                {
                    max = i;
                }
            }

            return max;
        }
    }
}
