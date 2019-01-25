using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using FullyQualifiedVector = ExceLint.Vector.RichVector;
using RelativeVector = System.Tuple<int, int, int>;
using Score = System.Collections.Generic.KeyValuePair<AST.Address, double>;
using HypothesizedFixes = System.Collections.Generic.Dictionary<AST.Address, System.Collections.Generic.Dictionary<string, ExceLint.Countable>>;
using Microsoft.FSharp.Core;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Collections.Immutable;
using FastDependenceAnalysis;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Clusters = System.Collections.Immutable.ImmutableHashSet<System.Collections.Immutable.ImmutableHashSet<AST.Address>>;
using ROInvertedHistogram = System.Collections.Immutable.ImmutableDictionary<AST.Address, System.Tuple<string, ExceLint.Scope.SelectID, ExceLint.Countable>>;

namespace ExceLintUI
{
    public class AnalysisCancelled : Exception { }

    [ComVisible(true)]
    public struct Analysis
    {
        public bool hasRun;
        public Score[] scores;
        public bool ranOK;
        public int cutoff;
        public Graph dag;
        public ExceLint.ErrorModel model;
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class WorkbookState : StandardOleMarshalObject
    {
        #region CONSTANTS
        // e * 1000
        public readonly static long MAX_DURATION_IN_MS = 5L * 60L * 1000L;  // 5 minutes
        public readonly static System.Drawing.Color GREEN = System.Drawing.Color.Green;
        public readonly static bool IGNORE_PARSE_ERRORS = true;
        public readonly static bool USE_WEIGHTS = true;
        public readonly static bool CONSIDER_ALL_OUTPUTS = true;
        public readonly static string CACHEDIRPATH = System.IO.Path.GetTempPath();
        #endregion CONSTANTS

        #region DATASTRUCTURES
        private Excel.Application _app;
        private Excel.Workbook _workbook;
        private double _tool_significance = 0.05;
        private ColorDict _colors = new ColorDict();
        private HashSet<AST.Address> _output_highlights = new HashSet<AST.Address>();
        private HashSet<AST.Address> _audited = new HashSet<AST.Address>();
        private Analysis _analysis;
        private AST.Address _flagged_cell;
        private bool _debug_mode = false;
        private bool _dag_changed = false;

        public int currentFlag = 0;

        #endregion DATASTRUCTURES

        #region BUTTON_STATE
        private bool _button_Analyze_enabled = true;
        private bool _button_MarkAsOK_enabled = false;
        private bool _button_FixError_enabled = false;
        private bool _button_clearColoringButton_enabled = false;
        private bool _button_RegularityMap_enabled = true;
        private bool _button_EntropyRanking_enabled = true;
        private Dictionary<Worksheet, bool> _visualization_shown = new Dictionary<Worksheet, bool>();
        private Dictionary<Worksheet, bool> _custodes_shown = new Dictionary<Worksheet, bool>();
        #endregion BUTTON_STATE

        public WorkbookState(Excel.Application app, Workbook workbook)
        {
            _app = app;
            _workbook = workbook;
            _analysis.hasRun = false;
        }

        public string WorkbookName
        {
            get { return _workbook.Name; }
        }

        public string Path
        {
            get
            {
                return _workbook.Path.TrimEnd(System.IO.Path.DirectorySeparatorChar) +
                       System.IO.Path.DirectorySeparatorChar;
            }
        }

        public double toolSignificance
        {
            get { return _tool_significance; }
            set { _tool_significance = value; }
        }

        public bool Analyze_Enabled
        {
            get { return _button_Analyze_enabled; }
            set { _button_Analyze_enabled = value; }
        }

        public bool MarkAsOK_Enabled
        {
            get { return _button_MarkAsOK_enabled; }
            set { _button_MarkAsOK_enabled = value; }
        }

     
        public bool ClearColoringButton_Enabled
        {
            get { return _button_clearColoringButton_enabled; }
            set { _button_clearColoringButton_enabled = value; }
        }

        public bool Visualization_Hidden(Worksheet w)
        {
            return !(_visualization_shown.ContainsKey(w) && _visualization_shown[w]);
        }

        public bool CUSTODES_Hidden(Worksheet w)
        {
            return !(_custodes_shown.ContainsKey(w) && _custodes_shown[w]);
        }

        public bool DebugMode
        {
            get { return _debug_mode; }
            set { _debug_mode = value; }
        }

        public Analysis inProcessAnalysis(long max_duration_in_ms, ExceLint.FeatureConf config, Graph g, Progress p)
        {
            // sanity check
            if (g.NumFormulas == 0)
            {
                return new Analysis { scores = null, ranOK = false, cutoff = 0 };
            }
            else
            {
                // run analysis
                FSharpOption<ExceLint.ErrorModel> mopt;
                try
                {
                    mopt = ExceLint.ModelBuilder.analyze(_app, config, g, _tool_significance, p);
                } catch (ExceLint.CommonTypes.NoFormulasException e)
                {
                    System.Windows.Forms.MessageBox.Show(e.Message);
                    throw new AnalysisCancelled();
                }

                if (FSharpOption<ExceLint.ErrorModel>.get_IsNone(mopt))
                {
                    throw new AnalysisCancelled();
                }
                else
                {
                    var model = mopt.Value;
                    Score[] scores = model.ranking();
                    int cutoff = model.Cutoff;
                    return new Analysis { scores = scores, ranOK = true, cutoff = cutoff, model = model, hasRun = true, dag = g };
                }
            }
        }

        public ExceLint.EntropyModelBuilder2.EntropyModel2 NewEntropyModelForWorksheet2(Worksheet w, ExceLint.FeatureConf conf, Graph g, ProgBar pb)
        {
            // create
            return ExceLint.ModelBuilder.initEntropyModel2(_app, conf, g, Progress.NOPProgress());
        }

        public void DrawImmutableClusters(Clusters clusters, ROInvertedHistogram ih, Worksheet ws)
        {
            var hs = new HashSet<HashSet<AST.Address>>();
            foreach (var c in clusters)
            {
                var c2 = new HashSet<AST.Address>(c);
                hs.Add(c2);
            }
            DrawClustersWithHistogram(hs, ih, ws);
        }

        public void ClearAllColors(Worksheet ws)
        {
            var initial_state = _app.ScreenUpdating;
            _app.ScreenUpdating = false;

            foreach (Excel.Range cell in ws.UsedRange)
            {
                cell.Interior.ColorIndex = 0;
            }
            _app.ScreenUpdating = initial_state;
        }

        public void DrawClustersWithHistogram(HashSet<HashSet<AST.Address>> clusters, ROInvertedHistogram ih, Worksheet ws)
        {
            // Disable screen updating
            var initial_state = _app.ScreenUpdating;
            _app.ScreenUpdating = false;

            // clear colors
            ClearAllColors(ws);

            // init cluster color map
            ClusterColorer clusterColors = new ClusterColorer(clusters, 0, 360, 0, ih);

            // do we stumble across protected cells along the way?
            var protCells = new List<AST.Address>();

            // paint
            foreach (var cluster in clusters)
            {
                System.Drawing.Color c = clusterColors.GetColor(cluster);

                foreach (AST.Address addr in cluster)
                {
                    if (!paintColor(addr, c))
                    {
                        protCells.Add(addr);
                    }
                }
            }

            // warn user if we could not highlight something
            if (protCells.Count > 0)
            {
                var names = String.Join(", ", protCells.Select(c => c.A1Local()));
                System.Windows.Forms.MessageBox.Show("WARNING: This workbook contains the following protected cells that cannot be highlighted:\n\n" + names);
            }

            // Enable screen updating
            _app.ScreenUpdating = initial_state;
        }

        public void flagNext(Worksheet ws)
        {
            if (FSharpOption<ExceLint.CommonTypes.ProposedFix[]>.get_IsSome(_analysis.model.Fixes))
            {
                var fixes = _analysis.model.Fixes.Value;
                if (currentFlag <= _analysis.model.Cutoff && currentFlag < fixes.Length)
                {
                    // get fix
                    var fix = fixes[currentFlag];

                    // don't update screen until done
                    var initial_state = _app.ScreenUpdating;
                    _app.ScreenUpdating = false;

                    // restore colors
                    ClearAllColors(ws);

                    // paint source
                    foreach(AST.Address a in fix.Source)
                    {
                        paintRed(a, 1.0);
                    }

                    // paint target
                    foreach(AST.Address a in fix.Target)
                    {
                        paintColor(a, System.Drawing.Color.Green);
                    }

                    // don't update screen until done
                    _app.ScreenUpdating = initial_state;

                    // activate and center
                    activateAndCenterOn(fix.Source.First(), _app);

                    currentFlag++;
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("No fixes remain.");

                    restoreOutputColors();
                    setTool(active: false);
                    currentFlag = 0;
                }
            }
        }

        public void analyze(long max_duration_in_ms, ExceLint.FeatureConf config, Boolean forceDAGBuild, ProgBar pb)
        {
            var sw = new System.Diagnostics.Stopwatch();
            sw.Start();

            // disable screen updating during analysis to speed things up
            var initial_state = _app.ScreenUpdating;
            _app.ScreenUpdating = false;

            // Also disable alerts; e.g., Excel thinks that compute-bound plugins
            // are deadlocked and alerts the user.  ExceLint really is just compute-bound.
            _app.DisplayAlerts = false;

            // build data dependence graph
            try
            {
                // test
                Graph g = new Graph(_app, (Worksheet)_app.ActiveSheet);

                _analysis = inProcessAnalysis(max_duration_in_ms, config, g, Progress.NOPProgress());

                if (!_analysis.ranOK)
                {
                    System.Windows.Forms.MessageBox.Show("This spreadsheet contains no formulas.");
                    return;
                } 
                //else
                //{
                //    var output = String.Join("\n", _analysis.model.ranking());
                //    System.Windows.Forms.Clipboard.SetText(output);
                //    System.Windows.Forms.MessageBox.Show("look in clipboard");
                //}
            }
            catch (AST.ParseException e)
            {
                // UI cleanup repeated here since the throw
                // below will cause the finally clause to be skipped
                _app.DisplayAlerts = true;
                _app.ScreenUpdating = initial_state;

                throw e;
            }
            finally
            {
                // Re-enable alerts
                _app.DisplayAlerts = true;

                // Enable screen updating when we're done
                _app.ScreenUpdating = initial_state;
            }

            sw.Stop();
        }

        private void activateAndCenterOn(AST.Address cell, Excel.Application app)
        {
            // go to worksheet
            RibbonHelper.GetWorksheetByName(cell.A1Worksheet(), _workbook.Worksheets).Activate();

            // COM object
            var comobj = ParcelCOMShim.Address.GetCOMObject(cell, app);

            // if the sheet is hidden, unhide it
            if (comobj.Worksheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
            {
                comobj.Worksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            }

            // if the cell's row is hidden, unhide it
            if ((bool)comobj.Rows.Hidden)
            {
                comobj.Rows.Hidden = false;
            }

            // if the cell's column is hidden, unhide it
            if ((bool)comobj.Columns.Hidden)
            {
                comobj.Columns.Hidden = false;
            }

            // ensure that the cell is wide enough that we can actually see it
            widenIfNecessary(comobj, app);

            // make sure that the printable area is big enough to show the cell;
            // don't change the printable area if the used range is a single cell
            int ur_width = comobj.Worksheet.UsedRange.Columns.Count;
            int ur_height = comobj.Worksheet.UsedRange.Rows.Count;
            if (ur_width != 1 || ur_height != 1)
            {
                comobj.Worksheet.PageSetup.PrintArea = comobj.Worksheet.UsedRange.Address;
            }

            // center screen on cell
            var visible_columns = app.ActiveWindow.VisibleRange.Columns.Count;
            var visible_rows = app.ActiveWindow.VisibleRange.Rows.Count;
            app.Goto(comobj, true);
            app.ActiveWindow.SmallScroll(Type.Missing, visible_rows / 2, Type.Missing, visible_columns / 2);

            // select highlighted cell
            // center on highlighted cell
            comobj.Select();

        }

        private void widenIfNecessary(Excel.Range comobj, Excel.Application app)
        {
            app.ScreenUpdating = false;
            var width = Convert.ToInt32(comobj.ColumnWidth);
            var height = Convert.ToInt32(comobj.RowHeight);
            comobj.Columns.AutoFit();
            comobj.Rows.AutoFit();

            if (Convert.ToInt32(comobj.ColumnWidth) < width)
            {
                comobj.ColumnWidth = width;
            }

            if (Convert.ToInt32(comobj.RowHeight) < height)
            {
                comobj.RowHeight = height;
            }

            app.ScreenUpdating = true;
        }

        private enum ProtectionLevel
        {
            None,
            Workbook,
            Worksheet
        }

        private ProtectionLevel unProtect(AST.Address cell)
        {
            // get cell COM object
            var com = ParcelCOMShim.Address.GetCOMObject(cell, _app);

            // check workbook protection
            if (((Excel.Workbook)com.Worksheet.Parent).HasPassword)
            {
                return ProtectionLevel.Workbook;
            }
            else
            {
                // try to unprotect worksheet
                try
                {
                    com.Worksheet.Unprotect(string.Empty);
                }
                catch
                {
                    return ProtectionLevel.Worksheet;
                }
            }
            return ProtectionLevel.None;
        }

        public bool Protected
        {
            get
            {
                // check workbook protection
                return _workbook.HasPassword;
            }
        }

        public bool WorksheetProtected(Worksheet w)
        {
            // check workbook protection
            if (Protected)
            {
                return true;
            } else
            {
                // try to unprotect worksheet
                try
                {
                    w.Unprotect(string.Empty);
                    return false;
                }
                catch
                {
                    return true;
                }
            }
        }

        public bool paintColor(AST.Address cell, System.Drawing.Color c)
        {
            // check cell protection
            if (unProtect(cell) != ProtectionLevel.None)
            {
                return false;
            }

            // get cell COM object
            var com = ParcelCOMShim.Address.GetCOMObject(cell, _app);

            // highlight cell
            com.Interior.Color = c;

            return true;
        }

        public void paintRed(AST.Address cell, double intensity)
        {
            // generate color
            byte A = System.Drawing.Color.Red.A;
            byte R = System.Drawing.Color.Red.R;
            byte G = Convert.ToByte((1.0 - intensity) * 255);
            byte B = Convert.ToByte((1.0 - intensity) * 255);
            var c = System.Drawing.Color.FromArgb(A, R, G, B);

            // highlight
            paintColor(cell, c);
        }

        public void saveColors(Worksheet ws)
        {
            foreach(Excel.Range cell in ws.UsedRange)
            {
                var addr = ParcelCOMShim.Address.AddressFromCOMObject(cell, _workbook);
                _colors.saveColorAt(
                    addr,
                    new CellColor { ColorIndex = (int)cell.Interior.ColorIndex, Color = (double)cell.Interior.Color }
                );
            }
        }

        public void restoreOutputColors()
        {
            var initial_state = _app.ScreenUpdating;
            _app.ScreenUpdating = false;

            if (_workbook != null)
            {
                foreach (KeyValuePair<AST.Address, CellColor> pair in _colors.all())
                {
                    var com = ParcelCOMShim.Address.GetCOMObject(pair.Key, _app);
                    com.Interior.ColorIndex = pair.Value.ColorIndex;
                }
                _colors.Clear();
            }
            _output_highlights.Clear();
            _colors.Clear();

            _app.ScreenUpdating = initial_state;
        }

        public void resetTool()
        {
            restoreOutputColors();
            _audited.Clear();
            currentFlag = 0;
            _analysis.hasRun = false;
            _custodes_shown[(Worksheet)_workbook.ActiveSheet] = false;
            setTool(active: false);
        }

        public void setTool(bool active)
        {
            _button_MarkAsOK_enabled = active;
            _button_FixError_enabled = active;
            _button_clearColoringButton_enabled = active;
            _button_Analyze_enabled = !active;
            _button_RegularityMap_enabled = !active;
            _button_EntropyRanking_enabled = !active;
        }

        internal void markAsOK(bool showFixes, Worksheet ws)
        {
            // the user told us that the cell was OK
            _audited.Add(_flagged_cell);

            // flag another value
            flagNext(ws);
        }
    }
}
