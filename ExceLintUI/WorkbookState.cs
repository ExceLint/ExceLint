using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Depends;
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
        //private double _tool_proportion = 0.05;
        private double _tool_significance = 0.05;
        private Dictionary<AST.Address, CellColor> _colors;
        private HashSet<AST.Address> _tool_highlights = new HashSet<AST.Address>();
        private HashSet<AST.Address> _output_highlights = new HashSet<AST.Address>();
        private HashSet<AST.Address> _known_good = new HashSet<AST.Address>();
        private Score[] _flaggable;
        private AST.Address _flagged_cell;
        private DAG _dag;
        private bool _debug_mode = false;
        private bool _dag_changed = false;
        private struct Analysis
        {
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
        #endregion BUTTON_STATE

        public WorkbookState(Excel.Application app, Excel.Workbook workbook)
        {
            _app = app;
            _workbook = workbook;
            _colors = new Dictionary<AST.Address, CellColor>();
        }

        public void DAGChanged()
        {
            _dag_changed = true;
        }

        //public double toolProportion
        //{
        //    get { return _tool_proportion; }
        //    set { _tool_proportion = value; }
        //}

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
            var cursor = _app.Selection;
            AST.Address cursorAddr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, _app.ActiveWorkbook);
            var cursorStr = "(" + cursorAddr.X + "," + cursorAddr.Y + ")";  // for sanity-preservation purposes

            // build DAG
            UpdateDAG(forceDAGBuild);

            Func<Progress, ExceLint.ErrorModel> f = (Progress p) =>
             {
                // find all vectors for formula under the cursor
                return new ExceLint.ErrorModel(config, _dag, _tool_significance, p);
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

        private delegate RelativeVector[] VectorSelector(AST.Address addr, DAG dag);
        private delegate FullyQualifiedVector[] AbsVectorSelector(AST.Address addr, DAG dag);

        private void getRawVectors(AbsVectorSelector f, Boolean forceDAGBuild)
        {
            // Disable screen updating during analysis to speed things up
            _app.ScreenUpdating = false;

            // build DAG
            UpdateDAG(forceDAGBuild);

            // get cursor location
            var cursor = _app.Selection;
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
            var cursor = _app.Selection;
            AST.Address cursorAddr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, _app.ActiveWorkbook);
            var cursorStr = "(" + cursorAddr.X + "," + cursorAddr.Y + ")";  // for sanity-preservation purposes

            // find all sources for formula under the cursor
            RelativeVector[] sourceVects = f(cursorAddr, _dag);

            // make string
            string[] sourceVectStrings = sourceVects.Select(vect => vect.ToString()).ToArray();
            var sourceVectsString = String.Join("\n", sourceVectStrings);

            // Enable screen updating when we're done
            _app.ScreenUpdating = true;

            System.Windows.Forms.MessageBox.Show("From: " + cursorStr + "\n\n" + sourceVectsString);
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
            VectorSelector f = (AST.Address addr, DAG dag) => ExceLint.Vector.getVectors(cell: addr, dag: dag, transitive: false, isForm: true, isRel: true, isMixed: true);
            getVectors(f, forceDAGBuild);
        }

        public void getFormulaRelVectors(Boolean forceDAGBuild)
        {
            VectorSelector f = (AST.Address addr, DAG dag) => ExceLint.Vector.getVectors(addr, dag, false, true, true, true);
            getVectors(f, forceDAGBuild);
        }

        public void getFormulaAbsVectors(Boolean forceDAGBuild)
        {
            VectorSelector f = (AST.Address addr, DAG dag) => ExceLint.Vector.getVectors(addr, dag, false, true, false, true);
            getVectors(f, forceDAGBuild);
        }

        public void getDataRelVectors(Boolean forceDAGBuild)
        {
            VectorSelector f = (AST.Address addr, DAG dag) => ExceLint.Vector.getVectors(addr, dag, false, false, true, true);
            getVectors(f, forceDAGBuild);
        }

        public void getDataAbsVectors(Boolean forceDAGBuild)
        {
            VectorSelector f = (AST.Address addr, DAG dag) => ExceLint.Vector.getVectors(addr, dag, false, false, false, true);
            getVectors(f, forceDAGBuild);
        }

        public void getRawFormulaVectors(Boolean forceDAGBuild)
        {
            AbsVectorSelector f = (AST.Address addr, DAG dag) => ExceLint.Vector.inputVectors(addr, dag, true);
            getRawVectors(f, forceDAGBuild);
        }

        public void getRawDataVectors(Boolean forceDAGBuild)
        {
            AbsVectorSelector f = (AST.Address addr, DAG dag) => ExceLint.Vector.outputVectors(addr, dag, true);
            getRawVectors(f, forceDAGBuild);
        }

        public void getL2NormSum(Boolean forceDAGBuild)
        {
            // Disable screen updating during analysis to speed things up
            _app.ScreenUpdating = false;

            // build DAG
            UpdateDAG(forceDAGBuild);

            // get cursor location
            var cursor = _app.Selection;
            AST.Address cursorAddr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, _app.ActiveWorkbook);
            var cursorStr = "(" + cursorAddr.X + "," + cursorAddr.Y + ")";  // for sanity-preservation purposes

            // find all sources for formula under the cursor
            double l2ns = ExceLint.Vector.DeepInputVectorRelativeL2NormSum.run(cursorAddr, _dag);

            // Enable screen updating when we're done
            _app.ScreenUpdating = true;

            System.Windows.Forms.MessageBox.Show(cursorStr + " = " + l2ns);
        }

        // this lets us reuse the progressbar for other work
        private T buildDAGAndDoStuff<T>(Boolean forceDAGBuild, Func<Progress,T> doStuff, long workMultiplier)
        {
            using (var pb = new ProgBar())
            {
                // create progress delegate
                ProgressBarIncrementer incr = () => pb.IncrementProgress();
                var p = new Progress(incr, workMultiplier);

                RefreshDAG(forceDAGBuild, p);

                return doStuff(p);
            }
        }

        private void UpdateDAG(Boolean forceDAGBuild)
        {
            Func<Progress,int> f = (Progress p) => 1;
            buildDAGAndDoStuff(forceDAGBuild, f, 1L);
        }

        private void RefreshDAG(Boolean forceDAGBuild, Progress p)
        {
            if (_dag == null)
            {
                _dag = DAG.DAGFromCache(forceDAGBuild, _app.ActiveWorkbook, _app, IGNORE_PARSE_ERRORS, CACHEDIRPATH, p);
            }
            else if (_dag_changed || forceDAGBuild)
            {
                _dag = new DAG(_app.ActiveWorkbook, _app, IGNORE_PARSE_ERRORS, p);
                _dag_changed = false;
                resetTool();
            }
        }

        public void analyze(long max_duration_in_ms, ExceLint.FeatureConf config, Boolean useHeatMap, Boolean forceDAGBuild)
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
                UpdateDAG(forceDAGBuild);

                Func<Progress, Analysis> f = (Progress p) =>
                {
                    // sanity check
                    if (_dag.getAllFormulaAddrs().Length == 0)
                    {
                        System.Windows.Forms.MessageBox.Show("This spreadsheet contains no formulas.");
                        _app.ScreenUpdating = true;
                        _flaggable = new KeyValuePair<AST.Address, double>[0];
                        return new Analysis { scores = null, ranOK = false, cutoff = 0 };
                    } else
                    {
                        // run analysis
                        var model = new ExceLint.ErrorModel(config, _dag, _tool_significance, p);
                        Score[] scores = model.rankByFeatureSum();
                        int cutoff = model.getSignificanceCutoff;
                        return new Analysis { scores = scores, ranOK = true, cutoff = cutoff, model = model };
                    }
                };

                var analysis = buildDAGAndDoStuff(forceDAGBuild, f, 3);

                if (!analysis.ranOK)
                {
                    return;
                }

                // assign scores to _flaggable
                _flaggable = analysis.scores;

                if (_flaggable.Length > 0 && useHeatMap)
                {
                    // calculate min/max heat map intensity
                    var min_score = analysis.scores[0].Value;
                    var max_score = analysis.scores[analysis.scores.Length - 1].Value;

                    // paint cells
                    foreach (Score s in analysis.scores)
                    {
                        // get score value
                        var sVal = s.Value;

                        // compute intensity
                        var intensity = 1.0;
                        if (max_score - min_score != 0)
                        {
                            intensity = (Convert.ToDouble(sVal - max_score) / Convert.ToDouble(min_score - max_score)) * 0.9 + 0.1;
                        }

                        // make it some shade of red
                        paintRed(s.Key, intensity);
                    }

                    setClearOnly();
                }

                // debug output
                if (_debug_mode && _flaggable.Length > 0)
                {
                    // scores
                    var score_str = String.Join("\n", _flaggable.Select(score => score.Key.A1FullyQualified() + " -> " + score.Value.ToString()));
                    if (score_str == "")
                    {
                        score_str = "empty";
                    }
                    System.Windows.Forms.Clipboard.SetText(score_str);
                    System.Windows.Forms.MessageBox.Show(score_str);

                    // time and space information
                    var time_str = "DAG construction ms: " + _dag.AnalysisMilliseconds + "\n" +
                                   "Feature scoring ms: " + analysis.model.ScoreTimeInMilliseconds + "\n" +
                                   "Num score entries: " + analysis.model.NumScoreEntries + "\n" +
                                   "Frequency counting ms: " + analysis.model.FrequencyTableTimeInMilliseconds + "\n" +
                                   "Num freq table entries: " + analysis.model.NumFreqEntries + "\n" +
                                   "Ranking ms: " + analysis.model.RankingTimeInMilliseconds + "\n" +
                                   "Total ranking length: " + analysis.model.NumRankedEntries;

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

            // center screen on cell
            var visible_columns = app.ActiveWindow.VisibleRange.Columns.Count;
            var visible_rows = app.ActiveWindow.VisibleRange.Rows.Count;
            app.Goto(comobj, true);
            app.ActiveWindow.SmallScroll(Type.Missing, visible_rows / 2, Type.Missing, visible_columns / 2);

            // select highlighted cell
            // center on highlighted cell
            comobj.Select();

        }

        public void flag()
        {
            //filter known_good
            _flaggable = _flaggable.Where(kvp => !_known_good.Contains(kvp.Key)).ToArray();
            if (_flaggable.Count() != 0)
            {
                // get TreeNode corresponding to most unusual score
                _flagged_cell = _flaggable.First().Key;
            }
            else
            {
                _flagged_cell = null;
            }

            if (_flagged_cell == null)
            {
                System.Windows.Forms.MessageBox.Show("No bugs remain.");
                resetTool();
            }
            else
            {
                // get cell COM object
                var com = ParcelCOMShim.Address.GetCOMObject(_flagged_cell, _app);

                // save old color
                var cc = new CellColor(com.Interior.ColorIndex, com.Interior.Color);
                if (!_colors.ContainsKey(_flagged_cell))
                {
                    _colors.Add(_flagged_cell, cc);
                }

                // highlight cell
                com.Interior.Color = System.Drawing.Color.Red;
                _tool_highlights.Add(_flagged_cell);

                // go to highlighted cell
                activateAndCenterOn(_flagged_cell, _app);

                // enable auditing buttons
                setTool(active: true);
            }
        }

        private void paintRed(AST.Address cell, double intensity)
        {
            // get cell COM object
            var com = ParcelCOMShim.Address.GetCOMObject(cell, _app);

            // save old color
            var cc = new CellColor(com.Interior.ColorIndex, com.Interior.Color);
            if (!_colors.ContainsKey(cell))
            {
                _colors.Add(cell, cc);
            }

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
            if (_workbook != null)
            {
                foreach (KeyValuePair<AST.Address, CellColor> pair in _colors)
                {
                    var com = ParcelCOMShim.Address.GetCOMObject(pair.Key, _app);
                    com.Interior.ColorIndex = pair.Value.ColorIndex;
                    com.Interior.Color = pair.Value.Color;
                }
                _colors.Clear();
            }
            _output_highlights.Clear();
            _colors.Clear();
        }

        public void resetTool()
        {
            restoreOutputColors();
            _known_good.Clear();
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

        internal void markAsOK()
        {
            // the user told us that the cell was OK
            _known_good.Add(_flagged_cell);

            // set the color of the cell to green
            var cell = ParcelCOMShim.Address.GetCOMObject(_flagged_cell, _app);
            cell.Interior.Color = GREEN;

            // restore output colors
            restoreOutputColors();

            // flag another value
            flag();
        }

        internal void fixError(Action<WorkbookState> setUIState, ExceLint.FeatureConf config)
        {
            var cell = ParcelCOMShim.Address.GetCOMObject(_flagged_cell, _app);
            // this callback gets run when the user clicks "OK"
            System.Action callback = () =>
            {
                // add the cell to the known good list
                _known_good.Add(_flagged_cell);

                // unflag the cell
                _flagged_cell = null;
                try
                {
                    // when a user fixes something, we need to re-run the analysis
                    analyze(MAX_DURATION_IN_MS, config, useHeatMap: false, forceDAGBuild: true);
                    // and flag again
                    flag();
                    // and then set the UI state
                    setUIState(this);
                }
                catch (Parcel.ParseException ex)
                {
                    System.Windows.Forms.Clipboard.SetText(ex.Message);
                    System.Windows.Forms.MessageBox.Show("Could not parse the formula string:\n" + ex.Message);
                    return;
                }
                catch (System.OutOfMemoryException ex)
                {
                    System.Windows.Forms.MessageBox.Show("Insufficient memory to perform analysis.");
                    return;
                }

            };
            // show the form
            var fixform = new CellFixForm(cell, GREEN, callback);
            fixform.Show();

            // restore output colors
            restoreOutputColors();
        }

        public string ToDOT()
        {
            var dag = new DAG(_app.ActiveWorkbook, _app, IGNORE_PARSE_ERRORS);
            return dag.ToDOT();
        }
    }
}
