using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using ExprOpt = Microsoft.FSharp.Core.FSharpOption<AST.Expression>;

namespace FastDependenceAnalysis
{
    public class Graph
    {
        private string _wsname;
        private string _wbname;
        private string _path;
        private string[][] _formulaTable;
        private object[][] _valueTable;
        private List<Dependence>[][] _dependenceTable;

        public Graph(Excel.Application a, Excel.Worksheet w)
        {
            // get names once
            _wsname = w.Name;
            _wbname = ((Excel.Workbook)w.Parent).Name;
            _path = ((Excel.Workbook)w.Parent).Path;

            // get used range
            Excel.Range urng = w.UsedRange;

            // formula table
            // invariant: null means not a formula
            _formulaTable = ReadFormulas(urng);

            // value table
            // invariant: null means empty cell
            _valueTable = ReadData(urng);

            // dependence table
            // invariant: table entry contains list of indices of dependency
            _dependenceTable = InitDependenceTable(_valueTable[0].Length, _valueTable.Length);

            // get dependence information from formulas
            for (int row = 0; row < _valueTable.Length; row++)
            {
                for (int col = 0; col < _valueTable[0].Length; col++)
                {
                    // is the cell a formula?
                    if (_formulaTable[row][col] != null)
                    {
                        // parse formula
                        ExprOpt astOpt = Parcel.parseFormula(_formulaTable[row][col], _path, _wbname, _wsname);
                        if (ExprOpt.get_IsSome(astOpt))
                        {
                            var ast = astOpt.Value;
                            // get range referencese
                            var rrefs = Parcel.rangeReferencesFromExpr(ast);
                            // get address references
                            var arefs = Parcel.addrReferencesFromExpr(ast);

                            // convert references into internal representation

                            // addresses first
                            for (int i = 0; i < arefs.Length; i++)
                            {
                                var addr = arefs[i];
                                // Excel row and column are 1-based
                                // subtract one to make them zero-based
                                _dependenceTable[row][col].Add(new Dependence(addr.WorksheetName != _wsname, addr.Row - 1, addr.Col - 1));
                            }

                            // references next
                            for (int i = 0; i < rrefs.Length; i++)
                            {
                                var rng = rrefs[i];
                                var addrs = rng.Addresses();
                                for (int j = 0; j < addrs.Length; j++)
                                {
                                    var addr = addrs[j];
                                    // Excel row and column are 1-based
                                    // subtract one to make them zero-based
                                    _dependenceTable[row][col].Add(new Dependence(addr.WorksheetName != _wsname, addr.Row - 1, addr.Col - 1));
                                }
                            }
                        }
                        else
                        {
                            // do ourselves a favor and remove entry from formula table
                            _formulaTable[row][col] = null;
                        }
                    }
                }
            }
        }

        private struct Dependence
        {
            public Dependence(bool onSheet, int row, int col)
            {
                OnSheet = onSheet;
                Row = row;
                Col = col;
            }

            public bool OnSheet { get; }

            public int Row { get; }

            public int Col { get; }
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

        private static List<Dependence>[][] InitDependenceTable(int width, int height)
        {
            var outer_y = new List<Dependence>[height][];
            for (int y = 0; y < height; y++)
            {
                outer_y[y] = new List<Dependence>[width];
                for (int x = 0; x < width; x++)
                {
                    outer_y[y][x] = new List<Dependence>();
                }
            }
            return outer_y;
        }


        private static AST.Address CellToAddress(int row, int col, string wsname, string wbname, string path)
        {
            return AST.Address.fromR1C1withMode(row, col, AST.AddressMode.Absolute, AST.AddressMode.Absolute, wsname, wbname, path);
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

        public AST.Address[] getAllFormulaAddrs()
        {
            var addrs = new HashSet<AST.Address>();
            for (int row = 0; row < _valueTable.Length; row++)
            {
                for (int col = 0; col < _valueTable[0].Length; col++)
                {
                    CellToAddress(row, col, _wsname, _wbname, _path);
                }
            }
            return addrs.ToArray();
        }
    }
}
