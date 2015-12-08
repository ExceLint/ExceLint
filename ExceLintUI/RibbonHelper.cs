using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Depends;
using Excel = Microsoft.Office.Interop.Excel;
using OptString = Microsoft.FSharp.Core.FSharpOption<string>;

namespace ExceLintUI
{
    public class CellColor
    {
        private int _colorindex;
        private double _color;

        public CellColor(int colorindex, double color)
        {
            _colorindex = colorindex;
            _color = color;
        }

        public double Color
        {
            get { return _color; }
        }

        public int ColorIndex
        {
            get { return _colorindex; }
        }
    }

    static class RibbonHelper
    {
        private static int TRANSPARENT_COLOR_INDEX = -4142;  //-4142 is the transparent default background

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
