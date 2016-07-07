using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using FullyQualifiedVector = ExceLint.Vector.FullyQualifiedVector;
using RelativeVector = System.Tuple<int, int, int>;
using Score = System.Collections.Generic.KeyValuePair<AST.Address, double>;

namespace ExceLintUI
{
    public class WorkbookState
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
        private HashSet<AST.Address> _tool_highlights = new HashSet<AST.Address>();
        private HashSet<AST.Address> _output_highlights = new HashSet<AST.Address>();
        private HashSet<AST.Address> _audited = new HashSet<AST.Address>();
        private Analysis _analysis;
        private AST.Address _flagged_cell;
        private Depends.DAG _dag;
        private bool _debug_mode = false;
        private bool _dag_changed = false;

        private struct Analysis
        {
            public bool hasRun;
            public Score[] scores;
            public bool ranOK;
            public int cutoff;
            public ExceLint.ErrorModel model;
        }
        #endregion DATASTRUCTURES

        #region BUTTON_STATE
        private bool _button_Analyze_enabled = true;
        private bool _button_MarkAsOK_enabled = false;
        private bool _button_FixError_enabled = false;
        private bool _button_clearColoringButton_enabled = false;
        private bool _button_showHeatMap_on = true;
        #endregion BUTTON_STATE

        public WorkbookState(Excel.Application app, Excel.Workbook workbook)
        {
            _app = app;
            _workbook = workbook;
            _analysis.hasRun = false;
        }

        public void DAGChanged()
        {
            _dag_changed = true;
        }

        public void ConfigChanged()
        {
            _analysis.hasRun = false;
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

        public bool FixError_Enabled
        {
            get { return _button_FixError_enabled; }
            set { _button_FixError_enabled = value; }
        }
        public bool ClearColoringButton_Enabled
        {
            get { return _button_clearColoringButton_enabled; }
            set { _button_clearColoringButton_enabled = value; }
        }

        public bool HeatMap_Hidden
        {
            get { return _button_showHeatMap_on; }
            set { _button_showHeatMap_on = value; }
        }

        public bool DebugMode
        {
            get { return _debug_mode; }
            set { _debug_mode = value; }
        }

        public void getSelected(ExceLint.FeatureConf config, Scope.Selector sel, Boolean forceDAGBuild)
        {
            // Disable screen updating during analysis to speed things up
            _app.ScreenUpdating = false;

            // get cursor location
            var cursor = (Excel.Range)_app.Selection;
            AST.Address cursorAddr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, _app.ActiveWorkbook);
            var cursorStr = "(" + cursorAddr.X + "," + cursorAddr.Y + ")";  // for sanity-preservation purposes

            // build DAG
            UpdateDAG(forceDAGBuild);

            Func<Depends.Progress, ExceLint.ErrorModel> f = (Depends.Progress p) =>
             {
                // find all vectors for formula under the cursor
                return new ExceLint.ErrorModel(_app, config, _dag, _tool_significance, p);
             };

            var model = buildDAGAndDoStuff(forceDAGBuild, f, 3);

            var output = model.inspectSelectorFor(cursorAddr, sel);

            // make output string
            string[] outputStrings = output.SelectMany(kvp => prettyPrintSelectScores(kvp)).ToArray();

            // Enable screen updating when we're done
            _app.ScreenUpdating = true;

            System.Windows.Forms.MessageBox.Show(cursorStr + "\n\n" + String.Join("\n", outputStrings));
        }

        private string[] prettyPrintSelectScores(KeyValuePair<AST.Address, Tuple<string, double>[]> addrScores)
        {
            var addr = addrScores.Key;
            var scores = addrScores.Value;

            return scores.Select(tup => addr + " -> " + tup.Item1 + ": " + tup.Item2).ToArray();
        }

        private delegate RelativeVector[] VectorSelector(AST.Address addr, Depends.DAG dag);
        private delegate FullyQualifiedVector[] AbsVectorSelector(AST.Address addr, Depends.DAG dag);

        private void getRawVectors(AbsVectorSelector f, Boolean forceDAGBuild)
        {
            // Disable screen updating during analysis to speed things up
            _app.ScreenUpdating = false;

            // build DAG
            UpdateDAG(forceDAGBuild);

            // get cursor location
            var cursor = (Excel.Range)_app.Selection;
            AST.Address cursorAddr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, _app.ActiveWorkbook);
            var cursorStr = "(" + cursorAddr.X + "," + cursorAddr.Y + ")";  // for sanity-preservation purposes

            // find all sources for formula under the cursor
            FullyQualifiedVector[] sourceVects = f(cursorAddr, _dag);

            // make string
            string[] sourceVectStrings = sourceVects.Select(vect => vect.ToString()).ToArray();
            var sourceVectsString = String.Join("\n", sourceVectStrings);

            // Enable screen updating when we're done
            _app.ScreenUpdating = true;

            System.Windows.Forms.MessageBox.Show("From: " + cursorStr + "\n\n" + sourceVectsString);
        }

        private void getVectors(VectorSelector f, Boolean forceDAGBuild)
        {
            // Disable screen updating during analysis to speed things up
            _app.ScreenUpdating = false;

            // build DAG
            UpdateDAG(forceDAGBuild);

            // get cursor location
            var cursor = (Excel.Range)_app.Selection;
            AST.Address cursorAddr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, _app.ActiveWorkbook);
            var cursorStr = "(" + cursorAddr.X + "," + cursorAddr.Y + ")";  // for sanity-preservation purposes

            // find all sources for formula under the cursor
            RelativeVector[] sourceVects = f(cursorAddr, _dag);

            // Enable screen updating when we're done
            _app.ScreenUpdating = true;

            if (sourceVects.Length == 0)
            {
                System.Windows.Forms.MessageBox.Show("No vectors.");
            } else
            {
                // make string
                string[] sourceVectStrings = sourceVects.Select(vect => vect.ToString()).ToArray();
                var sourceVectsString = String.Join("\n", sourceVectStrings);

                System.Windows.Forms.MessageBox.Show("From: " + cursorStr + "\n\n" + sourceVectsString);
            }
        }

        internal void SerializeDAG(Boolean forceDAGBuild)
        {
            if (_dag == null)
            {
                UpdateDAG(forceDAGBuild);
            }
            _dag.SerializeToDirectory(CACHEDIRPATH);
        }

        public void getMixedFormulaVectors(Boolean forceDAGBuild)
        {
            VectorSelector f = (AST.Address addr, Depends.DAG dag) => ExceLint.Vector.getVectors(cell: addr, dag: dag, transitive: false, isForm: true, isRel: true, isMixed: true, isOSI: true);
            getVectors(f, forceDAGBuild);
        }

        public void getFormulaRelVectors(Boolean forceDAGBuild)
        {
            VectorSelector f = (AST.Address addr, Depends.DAG dag) => ExceLint.Vector.getVectors(addr, dag, false, true, true, true, isOSI: true);
            getVectors(f, forceDAGBuild);
        }

        public void getFormulaAbsVectors(Boolean forceDAGBuild)
        {
            VectorSelector f = (AST.Address addr, Depends.DAG dag) => ExceLint.Vector.getVectors(addr, dag, false, true, false, true, isOSI: true);
            getVectors(f, forceDAGBuild);
        }

        public void getDataRelVectors(Boolean forceDAGBuild)
        {
            VectorSelector f = (AST.Address addr, Depends.DAG dag) => ExceLint.Vector.getVectors(addr, dag, false, false, true, true, isOSI: true);
            getVectors(f, forceDAGBuild);
        }

        public void getDataAbsVectors(Boolean forceDAGBuild)
        {
            VectorSelector f = (AST.Address addr, Depends.DAG dag) => ExceLint.Vector.getVectors(addr, dag, false, false, false, true, isOSI: true);
            getVectors(f, forceDAGBuild);
        }

        public void getRawFormulaVectors(Boolean forceDAGBuild)
        {
            AbsVectorSelector f = (AST.Address addr, Depends.DAG dag) => ExceLint.Vector.inputVectors(addr, dag, true);
            getRawVectors(f, forceDAGBuild);
        }

        public void getRawDataVectors(Boolean forceDAGBuild)
        {
            AbsVectorSelector f = (AST.Address addr, Depends.DAG dag) => ExceLint.Vector.outputVectors(addr, dag, true);
            getRawVectors(f, forceDAGBuild);
        }

        public void getL2NormSum(Boolean forceDAGBuild)
        {
            // Disable screen updating during analysis to speed things up
            _app.ScreenUpdating = false;

            // build DAG
            UpdateDAG(forceDAGBuild);

            // get cursor location
            var cursor = (Excel.Range)_app.Selection;
            AST.Address cursorAddr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, _app.ActiveWorkbook);
            var cursorStr = "(" + cursorAddr.X + "," + cursorAddr.Y + ")";  // for sanity-preservation purposes

            // find all sources for formula under the cursor
            double l2ns = ExceLint.Vector.DeepInputVectorRelativeL2NormSum.run(cursorAddr, _dag);

            // Enable screen updating when we're done
            _app.ScreenUpdating = true;

            System.Windows.Forms.MessageBox.Show(cursorStr + " = " + l2ns);
        }

        // this lets us reuse the progressbar for other work
        private T buildDAGAndDoStuff<T>(Boolean forceDAGBuild, Func<Depends.Progress,T> doStuff, long workMultiplier)
        {
            using (var pb = new ProgBar())
            {
                // create progress delegate
                Depends.ProgressBarIncrementer incr = () => pb.IncrementProgress();
                var p = new Depends.Progress(incr, workMultiplier);

                RefreshDAG(forceDAGBuild, p);

                return doStuff(p);
            }
        }

        private void UpdateDAG(Boolean forceDAGBuild)
        {
            Func<Depends.Progress,int> f = (Depends.Progress p) => 1;
            buildDAGAndDoStuff(forceDAGBuild, f, 1L);
        }

        private void RefreshDAG(Boolean forceDAGBuild, Depends.Progress p)
        {
            if (_dag == null)
            {
                _dag = Depends.DAG.DAGFromCache(forceDAGBuild, _app.ActiveWorkbook, _app, IGNORE_PARSE_ERRORS, CACHEDIRPATH, p);
            }
            else if (_dag_changed || forceDAGBuild)
            {
                _dag = new Depends.DAG(_app.ActiveWorkbook, _app, IGNORE_PARSE_ERRORS, p);
                _dag_changed = false;
                resetTool();
            }
        }

        public void toggleHeatMap(long max_duration_in_ms, ExceLint.FeatureConf config, Boolean forceDAGBuild)
        {
            if (HeatMap_Hidden)
            {
                if (!_analysis.hasRun)
                {
                    // run analysis
                    analyze(max_duration_in_ms, config, forceDAGBuild);
                }

                if (_analysis.cutoff > 0)
                {
                    // calculate min/max heat map intensity
                    var min_score = _analysis.scores[0].Value;
                    var max_score = _analysis.scores[_analysis.cutoff].Value;

                    // Disable screen updating 
                    _app.ScreenUpdating = false;

                    // paint cells
                    for (int i = 0; i <= _analysis.cutoff; i++)
                    {
                        var s = _analysis.scores[i];

                        // ensure that cell is unprotected or fail
                        if (unProtect(s.Key) != ProtectionLevel.None)
                        {
                            System.Windows.Forms.MessageBox.Show("Cannot highlight cell " + _flagged_cell.A1Local() + ". Cell is protected.");
                            return;
                        }

                        // make it some shade of red
                        paintRed(s.Key, intensity(min_score, max_score, s.Value));
                    }

                    // Enable screen updating
                    _app.ScreenUpdating = true;
                } else
                {
                    System.Windows.Forms.MessageBox.Show("No anomalies.");
                    return;
                }
            } else
            {
                restoreOutputColors();
            }
            toggleHeatMapSetting();
        }

        private double intensity(double min_score, double max_score, double score)
        {
            var lmax = Math.Log(max_score - min_score + 1);
            var lscore = Math.Log(score - min_score + 1);

            var shade = 1.0;
            if (lmax != lscore)
            {
                shade = (1 - (lscore / lmax)) / 2.0 + 0.5;
            }

            System.Diagnostics.Debug.Assert(shade >= 0.0 && shade <= 1.0);
            return shade;
        }

        public void analyze(long max_duration_in_ms, ExceLint.FeatureConf config, Boolean forceDAGBuild)
        {
            var sw = new System.Diagnostics.Stopwatch();
            sw.Start();

            // disable screen updating during analysis to speed things up
            _app.ScreenUpdating = false;

            // Also disable alerts; e.g., Excel thinks that compute-bound plugins
            // are deadlocked and alerts the user.  ExceLint really is just compute-bound.
            _app.DisplayAlerts = false;

            // build data dependence graph
            try
            {
                Func<Depends.Progress, Analysis> f = (Depends.Progress p) =>
                {
                    // sanity check
                    if (_dag.getAllFormulaAddrs().Length == 0)
                    {
                        System.Windows.Forms.MessageBox.Show("This spreadsheet contains no formulas.");
                        _app.ScreenUpdating = true;
                        return new Analysis { scores = null, ranOK = false, cutoff = 0 };
                    }
                    else
                    {
                        // run analysis
                        var model = new ExceLint.ErrorModel(_app, config, _dag, _tool_significance, p);
                        Score[] scores = model.rankByFeatureSum();
                        int cutoff = model.getSignificanceCutoff;
                        return new Analysis { scores = scores, ranOK = true, cutoff = cutoff, model = model, hasRun = true };
                    }
                };

                _analysis = buildDAGAndDoStuff(forceDAGBuild, f, 3);

                if (!_analysis.ranOK)
                {
                    return;
                }

                // debug output
                if (_debug_mode && _analysis.scores.Length > 0)
                {
                    // scores
                    var score_str = String.Join("\n", _analysis.scores.Select((score, idx) => {
                                        // prefix with cutoff marker, if applicable
                                        var prefix = "";
                                        if (idx == _analysis.cutoff + 1) { prefix = "--- CUTOFF ---\n"; }

                                        // enumerate causes
                                        var causes = _analysis.model.causeOf(score.Key);
                                        var causes_str = "\tcauses: [\n" + String.Join("\n", causes.Select(cause => "\t\t" + ExceLint.ErrorModel.prettyHistoBinDesc(cause.Key) + " = " + cause.Value)) + "\n\t]";

                        // print
                        return prefix + score.Key.A1FullyQualified() + " -> " + score.Value.ToString() + "\n" + causes_str + "\n\t" + "weight: " + _analysis.model.weightOf(score.Key);
                                    }));
                    if (score_str == "")
                    {
                        score_str = "empty";
                    }
                    System.Windows.Forms.Clipboard.SetText(score_str);
                    System.Windows.Forms.MessageBox.Show(score_str);

                    // time and space information
                    var time_str = "DAG construction ms: " + _dag.AnalysisMilliseconds + "\n" +
                                   "Feature scoring ms: " + _analysis.model.ScoreTimeInMilliseconds + "\n" +
                                   "Num score entries: " + _analysis.model.NumScoreEntries + "\n" +
                                   "Frequency counting ms: " + _analysis.model.FrequencyTableTimeInMilliseconds + "\n" +
                                   "Num freq table entries: " + _analysis.model.NumFreqEntries + "\n" +
                                   "Ranking ms: " + _analysis.model.RankingTimeInMilliseconds + "\n" +
                                   "Total ranking length: " + _analysis.model.NumRankedEntries;

                    System.Windows.Forms.Clipboard.SetText(time_str);
                    System.Windows.Forms.MessageBox.Show(time_str);

                }

                // Re-enable alerts
                _app.DisplayAlerts = true;

                // Enable screen updating when we're done
                _app.ScreenUpdating = true;
            }
            catch (Parcel.ParseException e)
            {
                // cleanup UI and then rethrow
                _app.ScreenUpdating = true;
                throw e;
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

            // make sure that the printable area is big enough to show the cell
            comobj.Worksheet.PageSetup.PrintArea = comobj.Worksheet.UsedRange.Address;

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

        public void flag()
        {
            // filter known_good & cut by cutoff index
            var flaggable = _analysis.scores
                                .Take(_analysis.cutoff + 1)
                                .Where(kvp => !_audited.Contains(kvp.Key)).ToArray();

            if (flaggable.Count() == 0)
            {
                System.Windows.Forms.MessageBox.Show("No remaining anomalies.");
                resetTool();
            }
            else
            {
                // get Score corresponding to most unusual score
                _flagged_cell = flaggable.First().Key;

                // ensure that cell is unprotected or fail
                if (unProtect(_flagged_cell) != ProtectionLevel.None)
                {
                    System.Windows.Forms.MessageBox.Show("Cannot highlight cell " + _flagged_cell.A1Local() + ". Cell is protected.");
                    return;
                } 

                // get cell COM object
                var com = ParcelCOMShim.Address.GetCOMObject(_flagged_cell, _app);

                // save old color
                _colors.saveColorAt(
                    _flagged_cell,
                    new CellColor { ColorIndex = (int)com.Interior.ColorIndex, Color = (double)com.Interior.Color }
                );

                // highlight cell
                com.Interior.Color = System.Drawing.Color.Red;
                _tool_highlights.Add(_flagged_cell);

                // go to highlighted cell
                activateAndCenterOn(_flagged_cell, _app);

                // enable auditing buttons
                setTool(active: true);
            }
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

        private void paintRed(AST.Address cell, double intensity)
        {
            // get cell COM object
            var com = ParcelCOMShim.Address.GetCOMObject(cell, _app);

            // save old color
            _colors.saveColorAt(
                cell,
                new CellColor { ColorIndex = (int)com.Interior.ColorIndex, Color = (double)com.Interior.Color }
            );

            // highlight cell
            byte A = System.Drawing.Color.Red.A;
            byte R = System.Drawing.Color.Red.R;
            byte G = Convert.ToByte((1.0 - intensity) * 255);
            byte B = Convert.ToByte((1.0 - intensity) * 255);
            var c = System.Drawing.Color.FromArgb(A, R, G, B);
            com.Interior.Color = c;
        }

        private void restoreOutputColors()
        {
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

            _app.ScreenUpdating = true;
        }

        public void resetTool()
        {
            restoreOutputColors();
            _audited.Clear();
            _analysis.hasRun = false;
            setTool(active: false);
        }

        private void setTool(bool active)
        {
            _button_MarkAsOK_enabled = active;
            _button_FixError_enabled = active;
            _button_clearColoringButton_enabled = active;
            _button_Analyze_enabled = !active;
        }

        private void setClearOnly()
        {
            _button_clearColoringButton_enabled = true;
        }

        private void toggleHeatMapSetting()
        {
            _button_showHeatMap_on = !_button_showHeatMap_on;
        }

        internal void markAsOK()
        {
            // the user told us that the cell was OK
            _audited.Add(_flagged_cell);

            // set the color of the cell to green
            var cell = ParcelCOMShim.Address.GetCOMObject(_flagged_cell, _app);
            cell.Interior.Color = GREEN;

            // restore output colors
            restoreOutputColors();

            // flag another value
            flag();
        }

        public string ToDOT()
        {
            var dag = new Depends.DAG(_app.ActiveWorkbook, _app, IGNORE_PARSE_ERRORS);
            return dag.ToDOT();
        }
    }
}
