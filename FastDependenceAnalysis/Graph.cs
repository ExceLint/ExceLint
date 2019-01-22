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
        private int _used_range_top;        // 1-based top y coordinate
        private int _used_range_bottom;     // 1-based bottom y coordinate
        private int _used_range_left;       // 1-based left-hand x coordinate
        private int _used_range_right;      // 1-based right-hand x coordinate
        private int _used_range_width;
        private int _used_range_height;

        public string Worksheet
        {
            get { return _wsname; }
        }
        public string Workbook
        {
            get { return _wbname; }
        }

        public string Path
        {
            get { return _path; }
        }

        public Dictionary<AST.Address, string> Values
        {
            get
            {
                var d = new Dictionary<AST.Address, string>();
                for (int row = 0; row < _valueTable.Length; row++)
                {
                    for (int col = 0; col < _valueTable[row].Length; col++)
                    {
                        var addr = CellToAddress(row, col, _wsname, _wbname, _path);
                        d.Add(addr, Convert.ToString(_valueTable[row][col]));
                    }
                }
                return d;
            }
        }

        public bool isOffSheet(AST.Address addr)
        {
            return addr.Path != _path || addr.WorkbookName != _wbname || addr.WorksheetName != _wsname;
        }

        public static string WorksheetName(Excel.Worksheet w)
        {
            return w.Name;
        }

        public static string WorkbookName(Excel.Worksheet w)
        {
            return ((Excel.Workbook)w.Parent).Name;
        }

        public static string WorkbookPath(Excel.Worksheet w)
        {
            var pthtmp = ((Excel.Workbook)w.Parent).Path;
            if (pthtmp.EndsWith(@"\"))
            {
                return pthtmp;
            }
            else
            {
                return pthtmp + @"\";
            }
        }

        public Graph(Excel.Application a, Excel.Worksheet w)
        {
            // get names once
            _wsname = WorksheetName(w);
            _wbname = WorkbookName(w);
            _path = WorkbookPath(w);

            // get used range
            Excel.Range urng = w.UsedRange;

            // get dimensions
            _used_range_left = urng.Column;                                  
            _used_range_right = urng.Columns.Count + _used_range_left - 1;   
            _used_range_top = urng.Row;                                      
            _used_range_bottom = urng.Rows.Count + _used_range_top - 1;
            _used_range_width = _used_range_right - _used_range_left + 1;
            _used_range_height = _used_range_bottom - _used_range_top + 1;

            // formula table
            // invariant: null means not a formula
            _formulaTable = ReadFormulas(urng, _used_range_left, _used_range_right, _used_range_top, _used_range_bottom, _used_range_width, _used_range_height);

            // value table
            // invariant: null means empty cell
            _valueTable = ReadData(urng, _used_range_left, _used_range_right, _used_range_top, _used_range_bottom, _used_range_width, _used_range_height);

            // dependence table
            // invariant: table entry contains list of indices of dependency
            _dependenceTable = InitDependenceTable(_valueTable[0].Length, _valueTable.Length);

            // get dependence information from formulas
            for (int row = 0; row < _formulaTable.Length; row++)
            {
                for (int col = 0; col < _formulaTable[row].Length; col++)
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
                                _dependenceTable[row][col].Add(new SingleDependence(isOffSheet(addr), addr.Row - 1, addr.Col - 1));
                            }

                            // references next
                            for (int i = 0; i < rrefs.Length; i++)
                            {
                                var rng = rrefs[i];
                                var addrs = rng.Addresses();
                                var sds = new List<SingleDependence>();

                                int maxCol = Int32.MinValue;
                                int minCol = Int32.MaxValue;
                                int maxRow = Int32.MinValue;
                                int minRow = Int32.MaxValue;
                                bool onSheet = true;

                                for (int j = 0; j < addrs.Length; j++)
                                {
                                    var addr = addrs[j];

                                    var addrOffSheet = isOffSheet(addr);

                                    // Excel row and column are 1-based
                                    // subtract one to make them zero-based
                                    var sd = new SingleDependence(addrOffSheet, addr.Row - 1, addr.Col - 1);
                                    sds.Add(sd);

                                    // find bounds
                                    if (addr.Row < minRow)
                                    {
                                        minRow = addr.Row;
                                    }

                                    if (addr.Row > maxRow)
                                    {
                                        maxRow = addr.Row;
                                    }

                                    if (addr.Col < minCol)
                                    {
                                        minCol = addr.Col;
                                    }

                                    if (addr.Col > maxCol)
                                    {
                                        maxCol = addr.Col;
                                    }

                                    // range is off-sheet if any of its addresses are off sheet
                                    if (onSheet && addrOffSheet)
                                    {
                                        onSheet = false;
                                    }
                                }

                                var rd = new RangeDependence(onSheet, minRow - 1, minCol - 1, maxRow - 1, maxCol - 1, sds);
                                _dependenceTable[row][col].Add(rd);
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

        interface Dependence
        {
            bool SingleRef { get; }
            bool OnSheet { get; }
        }

        private struct SingleDependence : Dependence
        {
            public SingleDependence(bool onSheet, int row, int col)
            {
                OnSheet = onSheet;
                Row = row;
                Col = col;
            }

            public bool OnSheet { get; }

            public int Row { get; }

            public int Col { get; }

            public bool SingleRef
            {
                get { return true; }
            }
        }

        private struct RangeDependence : Dependence
        {
            public RangeDependence(bool onSheet, int row_tl, int col_tl, int row_br, int col_br, List<SingleDependence> addrs)
            {
                OnSheet = onSheet;
                RowTop = row_tl;
                ColLeft = col_tl;
                RowBottom = row_br;
                ColRight = col_br;
                Addresses = addrs;
            }

            public bool OnSheet { get; }

            public int RowTop { get; }

            public int ColLeft { get; }

            public int RowBottom { get; }

            public int ColRight { get; }

            public List<SingleDependence> Addresses { get; }

            public bool SingleRef
            {
                get { return false; }
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

        private static SingleDependence AddressToCell(AST.Address addr, Graph g)
        {
            // if it's on-sheet, set flag to true
            return new SingleDependence(
                addr.Path == g.Path &&
                addr.WorkbookName == g.Workbook &&
                addr.WorksheetName == g.Worksheet,
                addr.Row - 1,
                addr.Col - 1);
        }

        private static object[][] ReadData(Excel.Range urng, int left, int right, int top, int bottom, int width, int height)
        {
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

                    // don't track if the cell contains nothing
                    if (data[y + 1, x + 1] != null) // adjust indices to be one-based
                    {
                        // copy the value in the cell
                        dataOutput[y][x] = data[y + 1, x + 1];
                    }

                    x_old = x;
                }
            }

            return dataOutput;
        }

        private static string[][] ReadFormulas(Excel.Range urng, int left, int right, int top, int bottom, int width, int height)
        {
            // init R1C1 extractor
            var regex = new Regex("^R([0-9]+)C([0-9]+)$", RegexOptions.Compiled);

            // init formula validator
            var fn_filter = new Regex("^=", RegexOptions.Compiled);

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
                            output[r - 1][c - 1] = f;
                        }
                    }
                }
            }
            return output;
        }

        public AST.Address[] getAllFormulaAddrs()
        {
            var addrs = new HashSet<AST.Address>();
            for (int row = 0; row < _formulaTable.Length; row++)
            {
                for (int col = 0; col < _formulaTable[row].Length; col++)
                {
                    if (_formulaTable[row][col] != null)
                    {
                        addrs.Add(CellToAddress(row, col, _wsname, _wbname, _path));
                    }
                }
            }
            return addrs.ToArray();
        }

        public bool isFormula(AST.Address addr)
        {
            // we arbitrarily say that off-sheet formulas are not formulas,
            // because there's no way to know
            var d = AddressToCell(addr, this);
            if (!d.OnSheet)
            {
                return false;
            }
            else
            {
                return _formulaTable[d.Row][d.Col] != null;
            }
        }

        public AST.Address[] allCells()
        {
            AST.Address[] output = new AST.Address[_valueTable.Length * _valueTable[0].Length];
            int i = 0;
            for (int row = 0; row < _valueTable.Length; row++)
            {
                for (int col = 0; col < _valueTable[row].Length; col++)
                {
                    var addr = CellToAddress(row, col, _wsname, _wbname, _path);
                    output[i] = addr;
                    i++;
                }
            }
            return output;
        }

        public AST.Address[] allCellsIncludingBlanks()
        {
            // unless I see a good reason not to do this, just run allCells()
            return allCells();
        }

        public int getPathClosureIndex(Tuple<string, string, string> path)
        {
            if (path.Item1 == _path && path.Item2 == _wbname && path.Item3 == _wsname)
            {
                return 0;
            }
            else
            {
                return 1;
            }
        }

        public string getFormulaAtAddress(AST.Address addr)
        {
            // if the formula is off-sheet return a constant function
            var d = AddressToCell(addr, this);
            if (d.OnSheet)
            {
                return _formulaTable[d.Row][d.Col];
            }
            else
            {
                return "=0";
            }
        }

        public string readCOMValueAtAddress(AST.Address addr)
        {
            // if the address is unresolvable or points to a
            // cell not in the dependence graph (e.g., empty cells)
            // return the empty string
            var d = AddressToCell(addr, this);
            if (!d.OnSheet)
            {
                return String.Empty;
            }
            else
            {
                try
                {
                    var s = Convert.ToString(_valueTable[d.Row][d.Col]);
                    if (s == null)
                    {
                        return String.Empty;
                    }
                    else
                    {
                        return s;
                    }
                }
                catch (Exception)
                {
                    return String.Empty;
                }

            }
        }

        public HashSet<AST.Address> getFormulaSingleCellInputs(AST.Address addr)
        {
            var d = AddressToCell(addr, this);
            var output = new HashSet<AST.Address>();
            if (d.OnSheet)
            {
                foreach (Dependence dabs in _dependenceTable[d.Row][d.Col])
                {
                    if (dabs.SingleRef)
                    {
                        var d2 = (SingleDependence)dabs;
                        var addr2 = CellToAddress(d2.Row, d2.Col, _wsname, _wbname, _path);
                        output.Add(addr2);
                    }
                }
            }

            return output;
        }

        public HashSet<AST.Range> getFormulaInputVectors(AST.Address faddr)
        {
            var d = AddressToCell(faddr, this);
            var output = new HashSet<AST.Range>();
            if (d.OnSheet)
            {
                foreach (Dependence dabs in _dependenceTable[d.Row][d.Col])
                {
                    if (!dabs.SingleRef)
                    {
                        var rd = (RangeDependence)dabs;
                        if (rd.OnSheet)
                        {
                            var topright = CellToAddress(rd.RowTop, rd.ColLeft, _wsname, _wbname, _path);
                            var bottomleft = CellToAddress(rd.RowBottom, rd.ColRight, _wsname, _wbname, _path);
                            AST.Range r = new AST.Range(topright, bottomleft);
                            output.Add(r);
                        }
                        else
                        {
                            // this dependence graph is lossy, so all we know is that it is
                            // not this sheet
                            string u = "unknown";
                            var topright = CellToAddress(rd.RowTop, rd.ColLeft, u, u, u);
                            var bottomleft = CellToAddress(rd.RowBottom, rd.ColRight, u, u, u);
                            AST.Range r = new AST.Range(topright, bottomleft);
                            output.Add(r);
                        }
                    }
                }
            }

            return output;
        }
    }
}
