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
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;

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
        public Depends.DAG dag;
        public ExceLint.ErrorModel model;
    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IWorkbookState
    {
        Analysis inProcessAnalysis(long max_duration_in_ms, ExceLint.FeatureConf config, Boolean forceDAGBuild, Depends.Progress p);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class WorkbookState : StandardOleMarshalObject, IWorkbookState, IDisposable
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
        private Depends.DAG _dag;
        private bool _debug_mode = false;
        private bool _dag_changed = false;
        private Dictionary<Worksheet, ExceLint.ClusterModelBuilder.ClusterModel> _m = new Dictionary<Worksheet, ExceLint.ClusterModelBuilder.ClusterModel>();
        // we deserialize cached dependence graphs in the background;
        // this lets us cancel deserialization in case the user quits
        // or if they explicitly request a new graph (like when they
        // request an analysis).
        private Object _dagLock = new Object();
        private DateTime _dagBuilt = DateTime.MinValue;
        private CancellationTokenSource _cts = new CancellationTokenSource();

        #endregion DATASTRUCTURES

        #region BUTTON_STATE
        private bool _button_Analyze_enabled = true;
        private bool _button_MarkAsOK_enabled = false;
        private bool _button_FixError_enabled = false;
        private bool _button_clearColoringButton_enabled = false;
        private Dictionary<Worksheet, bool> _visualization_shown = new Dictionary<Worksheet, bool>();
        private Dictionary<Worksheet, bool> _custodes_shown = new Dictionary<Worksheet, bool>();
        #endregion BUTTON_STATE

        public WorkbookState(Excel.Application app, Excel.Workbook workbook)
        {
            _app = app;
            _workbook = workbook;
            _analysis.hasRun = false;

            // if a cached DAG exists, load it eagerly
            if (Depends.DAG.CachedDAGExists(CACHEDIRPATH, _workbook.Name))
            {
                try
                {
                    Depends.DAG dag = null;

                    // grab a cancellation token
                    var token = _cts.Token;

                    // build the DAG in a background thread
                    var t = new Thread(() => {
                        var p = Depends.Progress.NOPProgress();
                        dag = Depends.DAG.DAGFromCache(false, _app.ActiveWorkbook, _app, IGNORE_PARSE_ERRORS, CACHEDIRPATH, p, token);
                    });
                    t.Start();

                    // check for cancellation
                    if (token.IsCancellationRequested)
                    {
                        return;
                    }

                    // write it
                    WriteDAG(dag);

                    // if the the cached DAG does not look like the
                    // spreadsheet we have open, update it
                    if (DAGChanged())
                    {
                        SerializeDAG(forceDAGBuild: true);
                    }
                } catch (Exception)
                {
                    // it's fine, do nothing
                }
            }
        }

        private void WriteDAG(Depends.DAG dag)
        {
            if (dag == null) return;

            lock (_dagLock)
            {
                if (_dagBuilt < dag.Built)
                {
                    _dag = dag;
                }
            }
        }

        public string WorkbookName
        {
            get { return _workbook.Name; }
        }

        public string Path
        {
            get { return _workbook.Path.TrimEnd(System.IO.Path.DirectorySeparatorChar) + System.IO.Path.DirectorySeparatorChar; }
        }

        public void MarkDAGAsChanged()
        {
            _dag_changed = true;
        }

        public bool DAGChanged()
        {
            // can't have changed if we've never built it before
            if (_dag == null)
            {
                return false;
            }
            return _dag.Changed(_workbook);
        }

        public void ConfigChanged()
        {
            _analysis.hasRun = false;
        }

        public FSharpOption<Analysis> getAnalysis()
        {
            if (_analysis.hasRun)
            {
                return FSharpOption<Analysis>.Some(_analysis);
            } else
            {
                return FSharpOption<Analysis>.None;
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

        public void getSelected(ExceLint.FeatureConf config, ExceLint.Scope.Selector sel, Boolean forceDAGBuild)
        {
            // Disable screen updating during analysis to speed things up
            _app.ScreenUpdating = false;

            // get cursor location
            var cursor = (Excel.Range)_app.Selection;
            AST.Address cursorAddr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, _app.ActiveWorkbook);
            var cursorStr = "(" + cursorAddr.X + "," + cursorAddr.Y + ")";  // for sanity-preservation purposes

            Func<Depends.Progress, ExceLint.ErrorModel> f = (Depends.Progress p) =>
             {
                 // find all vectors for formula under the cursor
                 //return new ExceLint.ErrorModel(_app, config, _dag, _tool_significance, p);

                 var mopt = ExceLint.ModelBuilder.analyze(_app, config, _dag, _tool_significance, p);
                 if (FSharpOption<ExceLint.ErrorModel>.get_IsNone(mopt))
                 {
                     throw new AnalysisCancelled();
                 } else {
                     return mopt.Value;
                 }
             };

            ExceLint.ErrorModel model;
            using (var pb = new ProgBar())
            {
                model = buildDAGAndDoStuff(forceDAGBuild, f, 3, pb);
            }

            var output = model.inspectSelectorFor(cursorAddr, sel, _dag);

            // make output string
            string[] outputStrings = output.SelectMany(kvp => prettyPrintSelectScores(kvp)).ToArray();

            // Enable screen updating when we're done
            _app.ScreenUpdating = true;

            System.Windows.Forms.MessageBox.Show(cursorStr + "\n\n" + String.Join("\n", outputStrings));
        }

        private string[] prettyPrintSelectScores(KeyValuePair<AST.Address, Tuple<string, ExceLint.Countable>[]> addrScores)
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
            using (var pb = new ProgBar())
            {
                UpdateDAG(forceDAGBuild, pb);
            }

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
            using (var pb = new ProgBar())
            {
                UpdateDAG(forceDAGBuild, pb);
            }

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
            using (var pb = new ProgBar())
            {
                UpdateDAG(forceDAGBuild, pb);
            }
            _dag.SerializeToDirectory(CACHEDIRPATH);
        }

        public void getL2NormSum(Boolean forceDAGBuild)
        {
            // Disable screen updating during analysis to speed things up
            _app.ScreenUpdating = false;

            // build DAG
            using (var pb = new ProgBar())
            {
                UpdateDAG(forceDAGBuild, pb);
            }

            // get cursor location
            var cursor = (Excel.Range)_app.Selection;
            AST.Address cursorAddr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, _app.ActiveWorkbook);
            var cursorStr = "(" + cursorAddr.X + "," + cursorAddr.Y + ")";  // for sanity-preservation purposes

            // find all sources for formula under the cursor
            ExceLint.Countable fc = ExceLint.Vector.DeepInputVectorRelativeL2NormSum.run(cursorAddr, _dag);
            double l2ns = 0.0;
            if (fc.IsNum)
            {
                var fcn = (ExceLint.Countable.Num)fc;
                l2ns = fcn.Item;
            }

            // Enable screen updating when we're done
            _app.ScreenUpdating = true;

            System.Windows.Forms.MessageBox.Show(cursorStr + " = " + l2ns);
        }

        private int EstimateNumberOfFormulas()
        {
            // simplistic formula validator
            var fn_filter = new Regex("^=", RegexOptions.Compiled);

            // count
            int i = 0;
            foreach (Worksheet w in _workbook.Worksheets)
            {
                // get used range
                var urng = w.UsedRange;

                // get dimensions
                var left = urng.Column;                      // 1-based left-hand y coordinate
                var right = urng.Columns.Count + left - 1;   // 1-based right-hand y coordinate
                var top = urng.Row;                          // 1-based top x coordinate
                var bottom = urng.Rows.Count + top - 1;      // 1-based bottom x coordinate

                // init
                int width = right - left + 1;
                int height = bottom - top + 1;

                if (left == right && top == bottom)
                {
                    var f = (string)urng.Formula;

                    if (fn_filter.IsMatch(f))
                    {
                        i++;
                    }
                }
                else
                {
                    // array read of formula cells
                    // note that this is a 1-based 2D multiarray
                    object[,] formulas = (object[,])urng.Formula;

                    // for every cell that is actually a formula, increment
                    for (int c = 1; c <= width; c++)
                    {
                        for (int r = 1; r <= height; r++)
                        {
                            var f = (string)formulas[r, c];
                            if (fn_filter.IsMatch(f))
                            {
                                i++;
                            }
                        }
                    }
                }
            }
            return i;
        }

        private T buildDAGAndDoStuff<T>(Boolean forceDAGBuild, Func<Depends.Progress, T> doStuff, long workMultiplier, ProgBar pb)
        {
            // create progress delegates
            Depends.ProgressBarIncrementer incr = n => pb.IncrementProgressN(n);
            Depends.ProgressBarReset reset = () => pb.Reset();
            var p = new Depends.Progress(incr, reset, workMultiplier);
            p.TotalWorkUnits = EstimateNumberOfFormulas();
            pb.registerCancelCallback(() => p.Cancel());

            RefreshDAG(forceDAGBuild, p);

            return doStuff(p);
        }

        private void UpdateDAG(Boolean forceDAGBuild, ProgBar pb)
        {
            Func<Depends.Progress,int> f = (Depends.Progress p) => 1;
            buildDAGAndDoStuff(forceDAGBuild, f, 1L, pb);
        }

        public bool DAGRefreshNeeded(Boolean forceDAGBuild)
        {
            return _dag == null || _dag_changed || forceDAGBuild;
        }

        private void RefreshDAG(Boolean forceDAGBuild, Depends.Progress p)
        {
            // cancel any currently-running DAG builds
            _cts.Cancel();

            // create new cancellation token source
            _cts = new CancellationTokenSource();

            if (_dag == null)
            {
                var dag = Depends.DAG.DAGFromCache(forceDAGBuild, _app.ActiveWorkbook, _app, IGNORE_PARSE_ERRORS, CACHEDIRPATH, p, _cts.Token);
                WriteDAG(dag);
            }
            else if (_dag_changed || forceDAGBuild)
            {
                var dag = new Depends.DAG(_app.ActiveWorkbook, _app, IGNORE_PARSE_ERRORS, p, DateTime.Now);
                WriteDAG(dag);
                _dag_changed = false;
                resetTool();
            } else
            {
                // manually bump progress bar so it looks to the user like we did something
                for (int i = 0; i < _dag.getAllFormulaAddrs().Length; i++)
                {
                    p.IncrementCounter();
                }
            }
        }

        public void toggleCUSTODES(string rootPath, string custodesPath, string javaPath, Workbook w)
        {
            // get current sheet
            var ws = (Worksheet)w.ActiveSheet;

            if (CUSTODES_Hidden(ws))
            {
                // get path to current spreadsheet
                var ssPath = System.IO.Path.Combine(w.Path, w.Name);
                var ssTempPath = System.IO.Path.Combine(
                    InstallScript.TempSpreadsheetDir(rootPath),
                    Globals.ThisAddIn.Application.ActiveWorkbook.Name);
                System.IO.File.Copy(ssPath, ssTempPath, overwrite: true);

                // run it
                var output = CUSTODES.getOutput(ssTempPath, custodesPath, javaPath);

                if (output.IsOKOutput)
                {
                    var ok = (CUSTODES.OutputResult.OKOutput)output;
                    var ok_output = ok.Item1;

                    // Inform user what is about to happen
                    System.Windows.Forms.MessageBox.Show("CUSTODES analysis complete.  Highlighting " + ok_output.Smells.Length + " cells.");


                    // Disable screen updating 
                    _app.ScreenUpdating = false;

                    // paint cells
                    for (int i = 0; i < ok_output.Smells.Length; i++)
                    {
                        var smell = ok_output.Smells[i];

                        // ensure that cell is unprotected or fail
                        if (unProtect(smell) != ProtectionLevel.None)
                        {
                            System.Windows.Forms.MessageBox.Show("Cannot highlight cell " + _flagged_cell.A1Local() + ". Cell is protected.");
                            return;
                        }

                        // make it bright red
                        paintRed(smell, 1.0);
                    }

                    // Enable screen updating
                    _app.ScreenUpdating = true;

                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Could not run CUSTODES.");
                    return;
                }
            } else
            {
                restoreOutputColors();
            }
            toggleCUSTODESSetting(ws);
        }

        public void toggleHeatMap(Worksheet w, long max_duration_in_ms, ExceLint.FeatureConf config, Boolean forceDAGBuild, ProgBar pb)
        {
            if (Visualization_Hidden(w))
            {
                if (!_analysis.hasRun)
                {
                    // run analysis
                    analyze(max_duration_in_ms, config, forceDAGBuild, pb);
                }

                if (_analysis.cutoff > 0)
                {
                    // calculate min/max heat map intensity
                    var max_score = _analysis.scores[0].Value;
                    var min_score = _analysis.scores[_analysis.cutoff].Value;

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
            toggleHeatMapSetting(w);
        }

        public void showSpectralPlot(long max_duration_in_ms, ExceLint.FeatureConf config, Boolean forceDAGBuild, ProgBar pb)
        {
            System.Windows.Forms.MessageBox.Show("disabled");
            //    if (!_analysis.hasRun)
            //    {
            //        // run analysis
            //        analyze(max_duration_in_ms, config, forceDAGBuild, pb);
            //    }

            //var plot = new SpectralPlot(_analysis.model);
            //plot.Show();
        }

        public void show3DScatterPlot(long max_duration_in_ms, ExceLint.FeatureConf config, Boolean forceDAGBuild, ProgBar pb)
        {
            System.Windows.Forms.MessageBox.Show("disabled");
            //if (!_analysis.hasRun)
            //{
            //    // run analysis
            //    analyze(max_duration_in_ms, config, forceDAGBuild, pb);
            //}

            //var plot = new Scatterplot3D(_analysis.model);
            //plot.Show();
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

        public Analysis inProcessAnalysis(long max_duration_in_ms, ExceLint.FeatureConf config, Boolean forceDAGBuild, Depends.Progress p)
        {
            // update if necessary
            RefreshDAG(forceDAGBuild, p);

            // sanity check
            if (_dag.getAllFormulaAddrs().Length == 0)
            {
                return new Analysis { scores = null, ranOK = false, cutoff = 0 };
            }
            else
            {
                // run analysis
                FSharpOption<ExceLint.ErrorModel> mopt;
                try
                {
                    mopt = ExceLint.ModelBuilder.analyze(_app, config, _dag, _tool_significance, p);
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
                    return new Analysis { scores = scores, ranOK = true, cutoff = cutoff, model = model, hasRun = true, dag = _dag };
                }
            }
        }

        public void LSHTest(ExceLint.FeatureConf conf, Boolean forceDAGBuild)
        {
            var p = Depends.Progress.NOPProgress();
            Excel.Application app = Globals.ThisAddIn.Application;

            // update if necessary
            RefreshDAG(forceDAGBuild, p);

            var m = ExceLint.ModelBuilder.VisualizeLSH(app, conf, _dag, 0.05, p);

            System.Windows.Forms.MessageBox.Show(m.ToGraphViz);
        }

        public Depends.DAG getDependenceGraph(Boolean forceDAGBuild)
        {
            // update if necessary
            var p = Depends.Progress.NOPProgress();
            RefreshDAG(forceDAGBuild, p);

            return _dag;
        }

        public void getLSHforAddr(AST.Address cursorAddr, Boolean forceDAGBuild)
        {
            // update if necessary
            var p = Depends.Progress.NOPProgress();
            RefreshDAG(forceDAGBuild, p);

            var resultant = ExceLint.Vector.ShallowInputVectorMixedFullCVectorResultantNotOSI.run(cursorAddr, _dag);
            var lsh = ExceLint.LSHCalc.h7(resultant);

            System.Windows.Forms.MessageBox.Show(lsh.MaskedBitsAsString(ExceLint.UInt128.MaxValue));
        }

        public double MoranForSelection(Excel.Range sel, Workbook wb, ExceLint.FeatureConf conf, Boolean forceDAGBuild)
        {
            // update if necessary
            var p = Depends.Progress.NOPProgress();
            RefreshDAG(forceDAGBuild, p);

            // convert range to set of points
            var rng = ParcelCOMShim.Range.RangeFromCOMObject(sel, wb);
            var addrs = rng.Addresses();

            // filter addresses by formulas
            var faddrs = addrs.Where(a => _dag.isFormula(a)).ToArray();

            // run feature on worksheet
            var scores = ExceLint.CommonFunctions.runEnabledFeatures(faddrs, _dag, conf, p);
            var flatScores = ExceLint.CommonFunctions.makeFlatScoreTable(scores);

            Func<AST.Address,double> x = addr => ExceLint.ClusterModelBuilder.X(addr, _dag, conf, flatScores);
            Func<AST.Address, AST.Address, double> w = (addr1, addr2) => ExceLint.ClusterModelBuilder.W(addr1, addr2);

            return ExceLint.ClusterModelBuilder.MoranCS(new HashSet<AST.Address>(addrs), x, w);
        }

        public void GetClusteringForWorksheet(Worksheet w, ExceLint.FeatureConf conf, Boolean forceDAGBuild, ProgBar pb)
        {
            Func<Depends.Progress, Unit> f = (p) =>
            {
                if (!_m.ContainsKey(w))
                {
                    Excel.Application app = Globals.ThisAddIn.Application;

                    // create
                    var m = ExceLint.ModelBuilder.initStepClusterModel(app, conf, _dag, 0.05, p);

                    // run agglomerative clustering until we reach inflection point
                    while (!m.NextStepIsKnee)
                    {
                        m.Step();

                        // debug visualize
                        restoreOutputColors();
                        DrawClusters(m.CurrentClustering);
                        System.Windows.Forms.MessageBox.Show("ok");
                    }

                    _m.Add(w, m);
                }
                else
                {
                    // fake progress bar if we've already done the work
                    for (int i = 0; i < _m[w].NumCells; i++)
                    {
                        pb.IncrementProgress();
                    }
                }
                return null;
            };

            // update DAG if necessary
            buildDAGAndDoStuff(forceDAGBuild, f, 3, pb);

            // draw
            DrawClusters(_m[w].CurrentClustering);
        }

        public void GetRegionsForWorksheet(Worksheet w, ExceLint.FeatureConf conf, Boolean forceDAGBuild, ProgBar pb)
        {
            Func<Depends.Progress, Unit> f = (p) =>
            {
                if (!_m.ContainsKey(w))
                {
                    Excel.Application app = Globals.ThisAddIn.Application;

                    // create
                    var m = ExceLint.ModelBuilder.initStepClusterModel(app, conf, _dag, 0.05, p);

                    // run agglomerative clustering until we reach inflection point
                    while (!m.NextStepIsKnee)
                    {
                        m.Step();
                    }

                    _m.Add(w, m);
                }
                else
                {
                    // fake progress bar if we've already done the work
                    for (int i = 0; i < _m[w].NumCells; i++)
                    {
                        pb.IncrementProgress();
                    }
                }
                return null;
            };

            // update DAG if necessary
            buildDAGAndDoStuff(forceDAGBuild, f, 3, pb);

            // extract regions
            var clustering = _m[w].Regions;

            // draw
            DrawClusters(clustering);
        }

        public void DrawClusters(HashSet<HashSet<AST.Address>> clusters)
        {
            // init cluster color map
            ClusterColorer clusterColors = new ClusterColorer(clusters, 0, 360, 0);

            // Disable screen updating
            _app.ScreenUpdating = false;

            // paint
            foreach (var cluster in clusters)
            {
                var c = clusterColors.GetColor(cluster);
                foreach (AST.Address addr in cluster)
                {
                    paintColor(addr, c);
                }
            }

            // Enable screen updating
            _app.ScreenUpdating = true;
        }

        public Analysis rawAnalysis(long max_duration_in_ms, ExceLint.FeatureConf config, Boolean forceDAGBuild, ProgBar pb)
        {
            Func<Depends.Progress, Analysis> f = (Depends.Progress p) =>
            {
                return inProcessAnalysis(max_duration_in_ms, config, forceDAGBuild, p);
            };

            return buildDAGAndDoStuff(forceDAGBuild, f, 3, pb);
        }

        public void analyze(long max_duration_in_ms, ExceLint.FeatureConf config, Boolean forceDAGBuild, ProgBar pb)
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
                _analysis = rawAnalysis(max_duration_in_ms, config, forceDAGBuild, pb);

                

                if (!_analysis.ranOK)
                {
                    System.Windows.Forms.MessageBox.Show("This spreadsheet contains no formulas.");
                    return;
                } else
                {
                    var output = String.Join("\n", _analysis.model.ranking());
                    System.Windows.Forms.Clipboard.SetText(output);
                    System.Windows.Forms.MessageBox.Show("look in clipboard");
                }
            }
            catch (AST.ParseException e)
            {
                // UI cleanup repeated here since the throw
                // below will cause the finally clause to be skipped
                _app.DisplayAlerts = true;
                _app.ScreenUpdating = true;

                throw e;
            }
            finally
            {
                // Re-enable alerts
                _app.DisplayAlerts = true;

                // Enable screen updating when we're done
                _app.ScreenUpdating = true;
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

        public static AST.Address[] hypothesizedFixes(AST.Address cell, ExceLint.ErrorModel model)
        {
            if (FSharpOption<HypothesizedFixes>.get_IsSome(model.Fixes))
            {
                var fixes = model.Fixes.Value[cell];
                return fixes.SelectMany(pair =>
                           model.Scores[pair.Key].Where(tup => tup.Item2 == pair.Value)
                       ).Select(tup => tup.Item1).ToArray();
            } else
            {
                return new AST.Address[] { };
            }
        }

        public void flag(bool showFixes)
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

                // go to highlighted cell
                activateAndCenterOn(_flagged_cell, _app);

                // enable auditing buttons
                setTool(active: true);

                // if this is COF, always show fixes
                if (_analysis.model.IsCOF)
                {
                    var fixes = _analysis.model.COFFixes[_flagged_cell].ToArray();

                    if (fixes.Length > 0)
                    {
                        var sb = new StringBuilder();

                        sb.AppendLine("ExceLint thinks that");
                        sb.AppendLine(_dag.getFormulaAtAddress(_flagged_cell));
                        sb.AppendLine("should look more like");

                        for (int i = 0; i < fixes.Length; i++)
                        {
                            // get formula at fix address
                            var f = _dag.getFormulaAtAddress(fixes[i]);
                            if (i > 0)
                            {
                                sb.Append("or ");
                            }
                            sb.AppendLine("address: " + fixes[i].A1Local().ToString() + ", formula: " + f);

                            // get cell COM object
                            var fix_com = ParcelCOMShim.Address.GetCOMObject(fixes[i], _app);

                            // save old color
                            _colors.saveColorAt(
                                fixes[i],
                                new CellColor { ColorIndex = (int)fix_com.Interior.ColorIndex, Color = (double)fix_com.Interior.Color }
                            );

                            // set color
                            fix_com.Interior.Color = System.Drawing.Color.Green;
                        }

                        System.Windows.Forms.MessageBox.Show(sb.ToString());
                    }
                }
                // if the user wants to see fixes, show them now
                else if (showFixes)
                {
                    var fixes = hypothesizedFixes(_flagged_cell, _analysis.model);
                    if (fixes.Length > 0)
                    {
                        var sb = new StringBuilder();

                        sb.AppendLine("ExceLint thinks that");
                        sb.AppendLine(_dag.getFormulaAtAddress(_flagged_cell));
                        sb.AppendLine("should look more like");

                        for (int i = 0; i < fixes.Length; i++)
                        {
                            // get formula at fix address
                            var f = _dag.getFormulaAtAddress(fixes[i]);
                            if (i > 0)
                            {
                                sb.Append("or ");
                            }
                            sb.AppendLine("address: " + fixes[i].A1Local().ToString() + ", formula: " + f);

                            // get cell COM object
                            var fix_com = ParcelCOMShim.Address.GetCOMObject(fixes[i], _app);

                            // save old color
                            _colors.saveColorAt(
                                fixes[i],
                                new CellColor { ColorIndex = (int)fix_com.Interior.ColorIndex, Color = (double)fix_com.Interior.Color }
                            );

                            // set color
                            fix_com.Interior.Color = System.Drawing.Color.Green;
                        }

                        System.Windows.Forms.MessageBox.Show(sb.ToString());
                    }
                }
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

        private void paintColor(AST.Address cell, System.Drawing.Color c)
        {
            // get cell COM object
            var com = ParcelCOMShim.Address.GetCOMObject(cell, _app);

            // save old color
            _colors.saveColorAt(
                cell,
                new CellColor { ColorIndex = (int)com.Interior.ColorIndex, Color = (double)com.Interior.Color }
            );

            // highlight cell
            com.Interior.Color = c;
        }

        private void paintRed(AST.Address cell, double intensity)
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

        public void toggleHeatMapSetting(Worksheet w)
        {
            if (!_visualization_shown.ContainsKey(w))
            {
                // if the worksheet is not in the dictionary,
                // it's because the visualization has never been on
                _visualization_shown.Add(w, true);
            } else
            {
                _visualization_shown[w] = !_visualization_shown[w];
            }
        }

        public void toggleCUSTODESSetting(Worksheet w)
        {
            if (!_custodes_shown.ContainsKey(w))
            {
                // if the worksheet is not in the dictionary,
                // it's because the visualization has never been on
                _custodes_shown.Add(w, true);
            }
            else
            {
                _custodes_shown[w] = !_custodes_shown[w];
            }
        }

        internal void markAsOK(bool showFixes)
        {
            // the user told us that the cell was OK
            _audited.Add(_flagged_cell);

            // set the color of the cell to green
            var cell = ParcelCOMShim.Address.GetCOMObject(_flagged_cell, _app);
            cell.Interior.Color = GREEN;

            // restore output colors
            restoreOutputColors();

            // flag another value
            flag(showFixes);
        }

        public string ToDOT()
        {
            var dag = new Depends.DAG(_app.ActiveWorkbook, _app, IGNORE_PARSE_ERRORS, DateTime.Now);
            return dag.ToDOT();
        }

        public string GetSquareMatrices(bool forceDAGBuild, bool normalizeRefSpace, bool normalizeSSSpace)
        {
            return null;
            //// Disable screen updating during analysis to speed things up
            //_app.ScreenUpdating = false;

            //// build DAG
            //using (var pb = new ProgBar())
            //{
            //    UpdateDAG(forceDAGBuild, pb);
            //}

            //// get cursor worksheet
            //var cursor = (Excel.Range)_app.Selection;
            //AST.Address cursorAddr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, _app.ActiveWorkbook);
            //var cWrksheet = cursorAddr.WorksheetName;

            //// get matrices for this sheet
            //var formulas = _dag.getAllFormulaAddrs().Where(c => c.WorksheetName == cWrksheet).ToArray();
            //var matrices = ExceLint.Vector.AllSquareVectors(formulas, _dag, normalizeRefSpace, normalizeSSSpace);

            //// convert to CSV
            //string[] rows = matrices.Select(v =>
            //{
            //    string[] s = new string[] { v.dx.ToString(), v.dy.ToString(), v.x.ToString(), v.y.ToString() };
            //    return String.Join(",", s);
            //}).ToArray();

            //string csv = String.Join("\n", rows);

            //_app.ScreenUpdating = true;

            //return csv;
        }

        public void Dispose()
        {
            _cts.Cancel();
        }
    }
}
