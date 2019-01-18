using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace FastDependenceAnalysis
{
    public class Graph
    {


        public Graph(Excel.Application a, Excel.Worksheet w)
        {
            // things to do:
            // 1. read in the CURRENT sheet
            // 2. allocate storage for all of the strings in the sheet in the usedrange
            // 3. allocate storage for all of the formulas in the sheet in the usedrange
            // 4. for each cell
            //   4.a. if that cell contains data, copy the string into a multi-array with the same coords
            //   4.b. if that cell contains a formula, copy the string into a multi-array with the same coords

        }

        public static void FastFormulaReadWorksheet(Excel.Application a, Excel.Worksheet w)
        {
            // get name once
            var wsname = w.Name;

            // get used range
            Excel.Range urng = w.UsedRange;

            // get formulas
            var formulas = ReadFormulas(urng);

            // get data
            var inputs = ReadData(urng);

            // init COM ref table (filled initially with null values)
            var refs = InitRangeTable(inputs.Length,inputs[0].Length);

            // process formulas
            foreach (var formula in formulas)
            {
                var addr = addrCache.getAddr(formula.Row, formula.Column, wsname, wbname, path);
                retVal.formulas.Add(addr, formula.Data);
                retVal.f2v.Add(addr, new HashSet<AST.Range>());
                retVal.f2i.Add(addr, new HashSet<AST.Address>());
            }

            // process data
            foreach (var input in inputs)
            {
                var addr = addrCache.getAddr(input.Row, input.Column, wsname, wbname, path);
                retVal.inputs.Add(addr, input.Data);
            }

            // process COM refs
            foreach (var drng in refs)
            {
                var addr = addrCache.getAddr(drng.Row, drng.Column, wsname, wbname, path);
                var formula = retVal.formulas.ContainsKey(addr) ? new Microsoft.FSharp.Core.FSharpOption<string>(retVal.formulas[addr]) : Microsoft.FSharp.Core.FSharpOption<string>.None;
                var cr = new ParcelCOMShim.LocalCOMRef(wb, worksheet, drng.Data, path, wbname, wsname, formula, 1, 1);
                retVal.allCells.Add(addr, cr);
            }
        }

        private static string[][] InitStringTable(int width, int height)
        {
            var outer_y = new string[height][];
            for (int y = 0; y < height; y++)
            {
                outer_y[y] = new string[width];
            }
            return outer_y;
        }

        private static object[][] InitObjectTable(int width, int height)
        {
            var outer_y = new object[height][];
            for (int y = 0; y < height; y++)
            {
                outer_y[y] = new object[width];
            }
            return outer_y;
        }

        private static object[][] InitRangeTable(int width, int height)
        {
            var outer_y = new Excel.Range[height][];
            for (int y = 0; y < height; y++)
            {
                outer_y[y] = new Excel.Range[width];
            }
            return outer_y;
        }

        private static object[][] ReadData(Excel.Range urng)
        {
            // get dimensions
            int left = urng.Column;                      // 1-based left-hand y coordinate
            int right = urng.Columns.Count + left - 1;   // 1-based right-hand y coordinate
            int top = urng.Row;                          // 1-based top x coordinate
            int bottom = urng.Rows.Count + top - 1;      // 1-based bottom x coordinate

            // init
            int width = right - left + 1;
            int height = bottom - top + 1;

            // output
            var dataOutput = InitObjectTable(width, height);

            // array read of data cells
            // note that this is a 1-based 2D multiarray
            // we grab this in array form so that we can avoid a COM
            // call for every blank-cell check
            object[,] data;
            // annoyingly, the return type for Value2 changes depending on the size of the range
            if (width == 1 && height == 1)
            {
                int[] lengths = new int[2] { 1, 1 };
                int[] lower_bounds = new int[2] { 1, 1 };
                data = (object[,])Array.CreateInstance(typeof(object), lengths, lower_bounds);
                data[1, 1] = urng.Value2;
            }
            else
            {   // ok, it really is an array
                data = (object[,])urng.Value2;
            }

            // if the worksheet contains nothing, data will be null
            if (data != null)
            {
                // for each COM object in the used range, create an address object
                // WITHOUT calling any methods on the COM object itself
                int x_old = -1;
                int x = -1;
                int y = 0;

                for (int i = 0; i < width * height; i++)
                {
                    // The basic idea here is that we know how Excel iterates over collections
                    // of cells.  The Excel.Range returned by UsedRange is always rectangular.
                    // Thus we can calculate the addresses of each COM cell reference without
                    // needing to incur the overhead of actually asking it for its address.
                    x = (x + 1) % width;
                    // increment y if x wrapped (x < x_old or x == x_old when width == 1)
                    y = x <= x_old ? y + 1 : y;

                    int c = x + left;
                    int r = y + top;

                    // don't track if the cell contains nothing
                    if (data[y + 1, x + 1] != null) // adjust indices to be one-based
                    {
                        // copy the value in the cell
                        dataOutput[r][c] = data[y + 1, x + 1];
                    }

                    x_old = x;
                }
            }

            return dataOutput;
        }

        private static string[][] ReadFormulas(Excel.Range urng)
        {
            // init R1C1 extractor
            var regex = new Regex("^R([0-9]+)C([0-9]+)$", RegexOptions.Compiled);

            // init formula validator
            var fn_filter = new Regex("^=", RegexOptions.Compiled);

            // get dimensions
            int left = urng.Column;                      // 1-based left-hand y coordinate
            int right = urng.Columns.Count + left - 1;   // 1-based right-hand y coordinate
            int top = urng.Row;                          // 1-based top x coordinate
            int bottom = urng.Rows.Count + top - 1;      // 1-based bottom x coordinate

            // init
            int width = right - left + 1;
            int height = bottom - top + 1;

            // output
            var output = InitStringTable(width, height);

            // if the used range is a single cell, Excel changes the type
            if (left == right && top == bottom)
            {
                var f = (string)urng.Formula;

                if (fn_filter.IsMatch(f))
                {
                    output[top][left] = f;
                }
            }
            else
            {
                // array read of formula cells
                // note that this is a 1-based 2D multiarray
                object[,] formulas = (object[,])urng.Formula;

                // for every cell that is actually a formula, add to 
                // formula dictionary & init formula lookup dictionaries
                for (int c = 1; c <= width; c++)
                {
                    for (int r = 1; r <= height; r++)
                    {
                        var f = (string)formulas[r, c];
                        if (fn_filter.IsMatch(f))
                        {
                            output[r + top - 1][c + left - 1] = f;
                        }
                    }
                }
            }
            return output;
        }
    }

    
}
