using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.FSharp.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Threading;
using ExceLintFileFormats;
using Application = Microsoft.Office.Interop.Excel.Application;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using System.Text;
using ExceLint;
using System.Collections.Immutable;
using Clusters = System.Collections.Immutable.ImmutableHashSet<System.Collections.Immutable.ImmutableHashSet<AST.Address>>;
using ROInvertedHistogram = System.Collections.Immutable.ImmutableDictionary<AST.Address, System.Tuple<string, ExceLint.Scope.SelectID, ExceLint.Countable>>;

namespace ExceLintUI
{
    public partial class ExceLintRibbon
    {
        private static bool USE_MULTITHREADED_UI = false;
        private static string DEFAULT_GROUND_TRUTH_FILENAME = "ground_truth";
        private static string JAVA_PATH = @"C:\ProgramData\Oracle\Java\javapath\java.exe";

        Dictionary<Excel.Workbook, WorkbookState> wbstates = new Dictionary<Excel.Workbook, WorkbookState>();
        System.Collections.Concurrent.ConcurrentDictionary<Excel.Workbook,Boolean> wbShutdown = new System.Collections.Concurrent.ConcurrentDictionary<Excel.Workbook,Boolean>();
        WorkbookState currentWorkbook;
        private ExceLintGroundTruth annotations;
        private string custodesPath = null;
        private AST.Address fixAddress = null; 
        private EntropyModelBuilder2.EntropyModel2 fixClusterModel = null;

        private string true_smells_csv = Properties.Settings.Default.CUSTODESTrueSmellsCSVPath;
        private string custodes_wbs_path = Properties.Settings.Default.CUSTODESWorkbooksPath;

        #region BUTTON_HANDLERS

        private void ClearEverything_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.restoreOutputColors();
        }

        private void LoadTrueSmells_Click(object sender, RibbonControlEventArgs e)
        {
            if (String.IsNullOrWhiteSpace(true_smells_csv))
            {
                var ofd = new System.Windows.Forms.OpenFileDialog();
                ofd.Title = "Where is the True Smells CSV?";
                ofd.ShowDialog();
                true_smells_csv = ofd.FileName;
                Properties.Settings.Default.CUSTODESTrueSmellsCSVPath = true_smells_csv;
                Properties.Settings.Default.Save();
            }
            
            if (String.IsNullOrWhiteSpace(custodes_wbs_path))
            {
                var odd = new System.Windows.Forms.FolderBrowserDialog();
                odd.Description = "Where are all the workbooks stored?";
                odd.ShowDialog();
                custodes_wbs_path = odd.SelectedPath;
                Properties.Settings.Default.CUSTODESWorkbooksPath = custodes_wbs_path;
                Properties.Settings.Default.Save();
            }

            // open & parse
            if (System.IO.Directory.Exists(custodes_wbs_path) && System.IO.File.Exists(true_smells_csv))
            {
                var allsmells = CUSTODES.GroundTruth.Load(custodes_wbs_path, true_smells_csv);

                // get true smells for this workbook
                var truesmells = allsmells.TrueSmellsbyWorkbook(Globals.ThisAddIn.Application.ActiveWorkbook.Name);
                var clustering = new HashSet<HashSet<AST.Address>>();
                clustering.Add(truesmells);

                // display
                currentWorkbook.DrawClusters(clustering);
            } else
            {
                System.Windows.Forms.MessageBox.Show("Can't find true smells CSV or workbook directory");
            }
        }

        private void cellIsFormula_Click(object sender, RibbonControlEventArgs e)
        {
            // get cursor location
            var cursor = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            AST.Address cursorAddr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, Globals.ThisAddIn.Application.ActiveWorkbook);

            // get config
            var conf = getConfig();

            // create progbar in main thread;
            // worker thread will call Dispose
            var pb = new ProgBar();

            // build the model
            Worksheet activeWs = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            var model = currentWorkbook.NewEntropyModelForWorksheet2(activeWs, getConfig(), this.forceBuildDAG.Checked, pb);

            // remove progress bar
            pb.Close();

            // display
            System.Windows.Forms.MessageBox.Show(cursorAddr.A1Local() + " = " + EntropyModelBuilder2.AddressIsFormulaValued(cursorAddr, model.InvertedHistogram, model.DependenceGraph).ToString());
        }

        private string ProposedFixesToString(CommonTypes.ProposedFix[] fixes)
        {
            // produce output string
            var sb = new StringBuilder();

            sb.Append("SOURCE");
            sb.Append(" -> ");
            sb.Append("TARGET");
            sb.Append(" = ");
            sb.Append("(");
            sb.Append(" NEG_INV_ENTROPY_DELTA ");
            sb.Append(" * ");
            sb.Append(" DOTPRODUCT ");
            sb.Append(")");
            sb.Append(" / ");
            sb.Append(" DISTANCE ");
            sb.Append(" = ");
            sb.Append(" RESULT ");
            sb.AppendLine();

            foreach (var fix in fixes)
            {
                // source
                var bbSource = Utils.BoundingRegion(fix.Source, 0);
                var sourceStart = bbSource.Item1.A1Local();
                var sourceEnd = bbSource.Item2.A1Local();
                sb.Append(sourceStart.ToString());
                sb.Append(":");
                sb.Append(sourceEnd.ToString());

                // separator
                sb.Append(" -> ");

                // target
                var bbTarget = Utils.BoundingRegion(fix.Target, 0);
                var targetStart = bbTarget.Item1.A1Local();
                var targetEnd = bbTarget.Item2.A1Local();
                sb.Append(targetStart.ToString());
                sb.Append(":");
                sb.Append(targetEnd.ToString());

                // separator
                sb.Append(" = ");

                // entropy * dp weight * inv_distance
                sb.Append("(");
                sb.Append(fix.E.ToString());
                sb.Append(" * ");
                sb.Append(fix.WeightedDotProduct.ToString());
                sb.Append(")");
                sb.Append(" / ");
                sb.Append(fix.Distance.ToString());
                sb.Append(" = ");
                sb.Append(fix.Score.ToString());

                // EOL
                sb.AppendLine();
            }

            return sb.ToString();
        }

        private void EntropyRanking_Click(object sender, RibbonControlEventArgs e)
        {
            // get dependence graph
            var graph = currentWorkbook.getDependenceGraph(false);

            // create progbar in main thread;
            // worker thread will call Dispose
            var pb = new ProgBar();

            // build the model
            Worksheet activeWs = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            var model = currentWorkbook.NewEntropyModelForWorksheet2(activeWs, getConfig(), this.forceBuildDAG.Checked, pb);

            // remove progress bar
            pb.Close();

            // get z for this worksheet
            var z = model.ZForWorksheet(activeWs.Name);

            // get fixes
            var fixes = model.Fixes(z);

            // get ranking
            var ranking = EntropyModelBuilder2.EntropyModel2.Ranking(fixes);

            // extract clusters
            var clusters = EntropyModelBuilder2.EntropyModel2.RankingToClusters(fixes);

            // draw
            currentWorkbook.DrawImmutableClusters(clusters, model.InvertedHistogram);

            // show message boxes
            System.Windows.Forms.MessageBox.Show("cutoff = " + model.Cutoff + "\n\nproposed fixes:\n" + ProposedFixesToString(fixes));
        }

        private void resetFixesButton_Click(object sender, RibbonControlEventArgs e)
        {
            // change button name
            FixClusterButton.Label = "Start Fix";

            // toss everything so that the user can do this again
            fixClusterModel = null;
            fixAddress = null;

            // redisplay visualiztion
            currentWorkbook.restoreOutputColors();
        }

        private Clusters ElideWhitespaceClusters(Clusters cs, ROInvertedHistogram ih, Depends.DAG graph)
        {
            var output = new HashSet<ImmutableHashSet<AST.Address>>();

            foreach (ImmutableHashSet<AST.Address> c in cs)
            {
                if (!c.All(a => EntropyModelBuilder.AddressIsWhitespaceValued(a, ih, graph)))
                {
                    output.Add(c);
                }
            }

            return output.ToImmutableHashSet();
        }

        private Clusters ElideStringClusters(Clusters cs, ROInvertedHistogram ih, Depends.DAG graph)
        {
            var output = new HashSet<ImmutableHashSet<AST.Address>>();

            foreach (ImmutableHashSet<AST.Address> c in cs)
            {
                if (!c.All(a => EntropyModelBuilder.AddressIsStringValued(a, ih, graph)))
                {
                    output.Add(c);
                }
            }

            return output.ToImmutableHashSet();
        }

        private Clusters PrettyClusters(Clusters cs, ROInvertedHistogram ih, Depends.DAG graph)
        {
            if (this.drawAllClusters.Checked)
            {
                return cs;
            } else
            {
                var cs1 = ElideStringClusters(cs, ih, graph);
                var cs2 = ElideWhitespaceClusters(cs1, ih, graph);
                return cs2;
            }
        }

        private string InvertedHistogramPrettyPrinter(ROInvertedHistogram ih)
        {
            var sb = new StringBuilder();

            sb.Append("[\n");
            foreach (var kvp in ih)
            {
                var c = kvp.Value.Item3;
                sb.Append(c.ToString());
                sb.Append("\n");
            }
            sb.Append("]\n");
            return sb.ToString();
        }

        private void FixClusterButton_Click(object sender, RibbonControlEventArgs e)
        {
            // get dependence graph
            var graph = currentWorkbook.getDependenceGraph(false);

            // get active sheet
            Worksheet activeWs = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            if (fixClusterModel == null)
            {
                // create progbar in main thread;
                // worker thread will call Dispose
                var pb = new ProgBar();

                // build the model
                var sw2 = System.Diagnostics.Stopwatch.StartNew();
                fixClusterModel = currentWorkbook.NewEntropyModelForWorksheet2(activeWs, getConfig(), this.forceBuildDAG.Checked, pb);

                // get z for worksheet
                var z = fixClusterModel.ZForWorksheet(activeWs.Name);

                // do visualization
                var histo2 = fixClusterModel.InvertedHistogram;
                var clusters2 = fixClusterModel.Clustering(z);
                sw2.Stop();

                var cl_filt2 = PrettyClusters(clusters2, histo2, graph);
                currentWorkbook.restoreOutputColors();
                currentWorkbook.DrawImmutableClusters(cl_filt2, fixClusterModel.InvertedHistogram);

                // remove progress bar
                pb.Close();

                // DEBUG: give me the graph
                //var gv = FasterBinaryMinEntropyTree.GraphViz(fixClusterModel.Trees[z]);
                //System.Windows.Forms.Clipboard.SetText(gv);
                //System.Windows.Forms.MessageBox.Show("graph in clipboard");
                //System.Windows.Forms.MessageBox.Show("score time ms: " + fixClusterModel.ScoreTimeMs);

                // change button name
                FixClusterButton.Label = "Select Source";
            }
            else
            {
                var app = Globals.ThisAddIn.Application;

                // get cursor location
                var cursor = (Excel.Range)app.Selection;

                // get address for cursor
                AST.Address cursorAddr =
                    ParcelCOMShim.Address.AddressFromCOMObject(cursor, Globals.ThisAddIn.Application.ActiveWorkbook);

                // get z for worksheet
                var z = fixClusterModel.ZForWorksheet(activeWs.Name);

                if (fixAddress == null)
                {
                    fixAddress = cursorAddr;

                    // change button name
                    FixClusterButton.Label = "Select Target";
                }
                else
                {
                    // do fix
                    var newModel = fixClusterModel.MergeCell(fixAddress, cursorAddr);

                    // compute change in entropy
                    var deltaE = fixClusterModel.EntropyDiff(z, newModel);

                    // save new model
                    fixClusterModel = newModel;

                    // redisplay visualiztion
                    var cl_filt = PrettyClusters(newModel.Clustering(z), newModel.InvertedHistogram, graph);
                    currentWorkbook.restoreOutputColors();
                    currentWorkbook.DrawImmutableClusters(cl_filt, fixClusterModel.InvertedHistogram);

                    // display output
                    System.Windows.Forms.MessageBox.Show("Change in entropy: " + deltaE);

                    // reset address
                    fixAddress = null;

                    // change button name
                    FixClusterButton.Label = "Select Source";
                }
            }
        }

        private void clusterForCell_Click(object sender, RibbonControlEventArgs e)
        {
            var ws = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            var app = Globals.ThisAddIn.Application;

            // get cursor location
            var cursor = (Excel.Range)app.Selection;

            // get address for cursor
            AST.Address cursorAddr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, Globals.ThisAddIn.Application.ActiveWorkbook);

            // create progbar in main thread;
            // worker thread will call Dispose
            var pb = new ProgBar();

            // build the model
            Worksheet activeWs = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            var model = currentWorkbook.NewEntropyModelForWorksheet2(activeWs, getConfig(), this.forceBuildDAG.Checked, pb);

            // get z for worksheet
            var z = model.ZForWorksheet(ws.Name);

            // get inverse lookup for clustering
            var addr2Cl = CommonFunctions.ReverseClusterLookup(model.Clustering(z));

            // get cluster for address
            var cluster = addr2Cl[cursorAddr];

            // remove progress bar
            pb.Close();

            // display cluster
            System.Windows.Forms.MessageBox.Show(String.Join(", ", cluster.Select(a => a.A1Local())));
        }

        private void nearestNeighborForCluster_Click(object sender, RibbonControlEventArgs e)
        {
            var w = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            var app = Globals.ThisAddIn.Application;

            // get cursor location
            var cursor = (Excel.Range)app.Selection;

            // get address for cursor
            AST.Address cursorAddr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, Globals.ThisAddIn.Application.ActiveWorkbook);

            // create progbar in main thread;
            // worker thread will call Dispose
            var pb = new ProgBar();

            // build the model
            Worksheet activeWs = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            var model = currentWorkbook.GetClusterModelForWorksheet(activeWs, getConfig(), this.forceBuildDAG.Checked, pb);

            // get inverse lookup for clustering
            var addr2Cl = CommonFunctions.ReverseClusterLookupMutable(model.CurrentClustering);

            // get cluster for address
            var cluster = addr2Cl[cursorAddr];

            // find the nearest neighbor
            var neighbor = model.NearestNeighborForCluster(cluster);

            // remove progress bar
            pb.Close();

            // display cluster
            System.Windows.Forms.MessageBox.Show(String.Join(", ", neighbor.Select(a => a.A1Local())));
        }

        private void RunCUSTODES_Click(object sender, RibbonControlEventArgs e)
        {
            string rootPath = null;

            // install CUSTODES if not already installed
            if (rootPath == null)
            {
                rootPath = InstallScript.InitDirs();
            }
            if (custodesPath == null)
            {
                custodesPath = InstallScript.InstallCUSTODES(rootPath);
            }

            // make sure that Excel does not think that
            // long-running analyses are deadlocks
            Globals.ThisAddIn.Application.DisplayAlerts = false;

            // run analysis and display on screen
            currentWorkbook.toggleCUSTODES(rootPath, custodesPath, JAVA_PATH, Globals.ThisAddIn.Application.ActiveWorkbook);

            setUIState(currentWorkbook);

            // reset alerts
            Globals.ThisAddIn.Application.DisplayAlerts = true;
        }

        private void LSHTest_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.LSHTest(getConfig(), this.forceBuildDAG.Checked);
        }

        private void getLSH_Click(object sender, RibbonControlEventArgs e)
        {
            // get cursor location
            var cursor = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            AST.Address cursorAddr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, Globals.ThisAddIn.Application.ActiveWorkbook);

            currentWorkbook.getLSHforAddr(cursorAddr, false);
        }

        private void VectorForCell_Click(object sender, RibbonControlEventArgs e)
        {
            // get cursor location
            var cursor = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            AST.Address cursorAddr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, Globals.ThisAddIn.Application.ActiveWorkbook);

            // get config
            var conf = getConfig();

            // get dependence graph
            var dag = currentWorkbook.getDependenceGraph(this.forceBuildDAG.Checked);

            var sb = new StringBuilder();

            // get vector for each enabled feature
            var feats = conf.EnabledFeatures;
            for (int i = 0; i < feats.Length; i++)
            {
                // run feature
                sb.Append(feats[i]);
                sb.Append(" = ");
                sb.Append(conf.get_FeatureByName(feats[i]).Invoke(cursorAddr).Invoke(dag).ToString());
                sb.Append("\n");
            }

            // display
            System.Windows.Forms.MessageBox.Show(sb.ToString());
        }

        public WorkbookState CurrentWorkbook { 
            get
            {
                return currentWorkbook;
            }
        }

        private void AnalyzeButton_Click(object sender, RibbonControlEventArgs e)
        {
            // check for debug checkbox
            currentWorkbook.DebugMode = this.DebugOutput.Checked;

            // get significance threshold
            var sig = getPercent(this.significanceTextBox.Text, this.significanceTextBox.Label);

            // workbook- and UI-update callback
            Action<WorkbookState> updateWorkbook = (WorkbookState wbs) =>
            {
                this.currentWorkbook = wbs;
                setUIState(currentWorkbook);
            };

            // create progbar in main thread;
            // worker thread will call Dispose
            var pb = new ProgBar();

            // call task in new thread and do not wait
            //Task t = Task.Run(() => DoAnalysis(sig, currentWorkbook, getConfig(), this.forceBuildDAG.Checked, updateWorkbook, pb));
            DoAnalysis(sig, currentWorkbook, getConfig(), this.forceBuildDAG.Checked, updateWorkbook, pb, showFixes.Checked);
        }

        public static void DoAnalysis(FSharpOption<double> sigThresh, WorkbookState wbs, ExceLint.FeatureConf conf, bool forceBuildDAG, Action<WorkbookState> updateState, ProgBar pb, bool showFixes)
        {
            if (sigThresh == FSharpOption<double>.None)
            {
                return;
            }
            else
            {
                wbs.toolSignificance = sigThresh.Value;
                try
                {
                    wbs.analyze(WorkbookState.MAX_DURATION_IN_MS, conf, forceBuildDAG, pb);
                    wbs.MarkAsOK_Enabled = true;
                    wbs.flagNext();
                    //wbs.flag(showFixes);
                    updateState(wbs);

                    // debug output
                    if (wbs.DebugMode)
                    {
                        var debug_info = prepareDebugInfo(wbs);
                        var timing_info = prepareTimingInfo(wbs);
                        RunInSTAThread(() => printDebugInfo(debug_info, timing_info));
                    }

                    pb.GoAway();
                }
                catch (AST.ParseException ex)
                {
                    RunInSTAThread(() =>
                    {
                        System.Windows.Forms.Clipboard.SetText(ex.Message);
                        System.Windows.Forms.MessageBox.Show("Could not parse the formula string:\n" + ex.Message);
                    });
                }
                catch (AnalysisCancelled)
                {
                    RunInSTAThread(() =>
                    {
                        System.Windows.Forms.MessageBox.Show("Analysis cancelled.");
                    });
                }
                catch (OutOfMemoryException)
                {
                    RunInSTAThread(() =>
                    {
                        System.Windows.Forms.MessageBox.Show("Insufficient memory to perform analysis.");
                    });
                }
                catch (Exception ex)
                {
                    RunInSTAThread(() =>
                    {
                        var msg = "Runtime exception. This message has been copied to your clipboard.\n" + ex.Message + "\n\nStack trace:\n" + ex.StackTrace;
                        System.Windows.Forms.Clipboard.SetText(msg);
                        System.Windows.Forms.MessageBox.Show(msg);
                    });
                }
            }
        }

        private static string prepareDebugInfo(WorkbookState wbs)
        {
            var a = wbs.getAnalysis();

            if (FSharpOption<Analysis>.get_IsNone(a))
            {
                return "";
            }

            var analysis = a.Value;

            if (analysis.scores.Length == 0)
            {
                return "";
            }

            // scores
            var score_str = String.Join("\n", analysis.scores.Select((score, idx) => {
                // prefix with cutoff marker, if applicable
                var prefix = "";
                if (idx == analysis.cutoff + 1) { prefix = "--- CUTOFF ---\n"; }

                // enumerate causes
                string causes_str = "";
                try
                {
                    var causes = analysis.model.causeOf(score.Key);
                    causes_str = "\tcauses: [\n" +
                                     String.Join("\n", causes.Select(cause => {
                                         var causeScore = cause.Value.Item1;
                                         var causeWeight = cause.Value.Item2;
                                         return "\t\t" + ExceLint.ErrorModel.prettyHistoBinDesc(cause.Key) + ": (CSS weight) * score = " + causeWeight + " x " + causeScore + " = " + causeWeight * causeScore;
                                     })) + "\n\t]";
                } catch (Exception) { }

                // print
                return prefix + score.Key.A1FullyQualified() + " -> " + score.Value.ToString() + "\n" + causes_str + "\n\t" + "intrinsic anomalousness weight: " + analysis.model.weightOf(score.Key);
            }));
            if (score_str == "")
            {
                score_str = "empty";
            }

            return score_str;
        }

        private static string prepareTimingInfo(WorkbookState wbs)
        {
            var a = wbs.getAnalysis();

            if (FSharpOption<Analysis>.get_IsNone(a))
            {
                return "";
            }

            var analysis = a.Value;

            // time and space information
            var time_str = "Marshaling ms: " + analysis.dag.TimeMarshalingMilliseconds + "\n" +
                           "Parsing ms: " + analysis.dag.TimeParsingMilliseconds + "\n" +
                           "Graph construction ms: " + analysis.dag.TimeGraphConstructionMilliseconds + "\n" +
                           "Feature scoring ms: " + analysis.model.ScoreTimeInMilliseconds + "\n" +
                           "Num score entries: " + analysis.model.NumScoreEntries + "\n" +
                           "Frequency counting ms: " + analysis.model.FrequencyTableTimeInMilliseconds + "\n" +
                           "Conditioning set size ms: " + analysis.model.ConditioningSetSizeTimeInMilliseconds + "\n" +
                           "Causes ms: " + analysis.model.CausesTimeInMilliseconds + "\n" +
                           "Num freq table entries: " + analysis.model.NumFreqEntries + "\n" +
                           "Ranking ms: " + analysis.model.RankingTimeInMilliseconds + "\n" +
                           "Total ranking length: " + analysis.model.NumRankedEntries;

            return time_str;
        }

        private static void printDebugInfo(string debug_info, string time_info)
        {
            if (!String.IsNullOrEmpty(debug_info))
            {
                var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                var debugInfoPath = System.IO.Path.Combine(desktopPath, "ExceLintDebugInfo.txt");

                System.IO.File.WriteAllText(debugInfoPath, debug_info);

                System.Windows.Forms.MessageBox.Show("Debug information written to file:\n" + debugInfoPath);
            }

            System.Windows.Forms.MessageBox.Show(time_info);
        }

        private static void RunInSTAThread(ThreadStart t)
        {
            if (USE_MULTITHREADED_UI)
            {
                Thread thread = new Thread(t);
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();
            } else
            {
                t.Invoke();
            }
        }

        private void MarkAsOKButton_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.markAsOK(showFixes.Checked);
            setUIState(currentWorkbook);
        }

        private void StartOverButton_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.resetTool();
            setUIState(currentWorkbook);
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.getL2NormSum(forceDAGBuild: forceBuildDAG.Checked);
        }

        private void ToDOT_Click(object sender, RibbonControlEventArgs e)
        {
            var graphviz = currentWorkbook.ToDOT();
            System.Windows.Clipboard.SetText(graphviz);
            System.Windows.Forms.MessageBox.Show("Done. Graph is in the clipboard.");
        }

        private void colSelect_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.getSelected(getConfig(), ExceLint.Scope.Selector.SameColumn, forceDAGBuild: forceBuildDAG.Checked);
        }

        private void rowSelected_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.getSelected(getConfig(), ExceLint.Scope.Selector.SameRow, forceDAGBuild: forceBuildDAG.Checked);
        }

        private void showHeatmap_Click(object sender, RibbonControlEventArgs e)
        {
            var w = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            if (currentWorkbook.Visualization_Hidden(w))
            {
                // create progbar in main thread;
                // worker thread will call Dispose
                var pb = new ProgBar();

                // show a cluster visualization
                Worksheet activeWs = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                currentWorkbook.GetClusteringForWorksheet(activeWs, getConfig(), this.forceBuildDAG.Checked, pb);

                // remove progress bar
                pb.Close();
            }
            else
            {
                // erase the cluster visualization
                currentWorkbook.resetTool();
            }

            // toggle button
            currentWorkbook.toggleHeatMapSetting(w);

            // set UI state
            setUIState(currentWorkbook);
        }

        private static void DoHeatmap(FSharpOption<double> sigThresh, WorkbookState wbs, ExceLint.FeatureConf conf, bool forceBuildDAG, Action<WorkbookState> updateState, ProgBar pb)
        {
            var w = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            if (sigThresh == FSharpOption<double>.None)
            {
                return;
            }
            else
            {
                wbs.toolSignificance = sigThresh.Value;
                try
                {
                    // if, BEFORE the analysis, the user requests debug info
                    // AND the heatmap is PRESENTLY HIDDEN, show the debug info
                    // whether there ends up being a heatmap to show or not.
                    var debug_display = wbs.DebugMode && wbs.Visualization_Hidden(w);

                    wbs.toggleHeatMap(w, WorkbookState.MAX_DURATION_IN_MS, conf, forceBuildDAG, pb);
                    updateState(wbs);

                    // debug output
                    if (debug_display)
                    {
                        var debug_info = prepareDebugInfo(wbs);
                        var timing_info = prepareTimingInfo(wbs);
                        RunInSTAThread(() => printDebugInfo(debug_info, timing_info));
                    }

                    pb.GoAway();
                }
                catch (AST.ParseException ex)
                {
                    RunInSTAThread(() =>
                    {
                        System.Windows.Forms.Clipboard.SetText(ex.Message);
                        System.Windows.Forms.MessageBox.Show("Could not parse the formula string:\n" + ex.Message);
                    });
                }
                catch (AnalysisCancelled)
                {
                    RunInSTAThread(() =>
                    {
                        System.Windows.Forms.MessageBox.Show("Analysis cancelled.");
                    });
                }
                catch (OutOfMemoryException)
                {
                    RunInSTAThread(() =>
                    {
                        System.Windows.Forms.MessageBox.Show("Insufficient memory to perform analysis.");
                    });
                }
            }
        }

        private void allCellsFreq_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void columnCellsFreq_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void rowCellsFreq_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void DebugOutput_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void forceBuildDAG_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void inferAddrModes_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void allCells_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void weightByIntrinsicAnomalousness_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void inDegree_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void outDegree_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void combinedDegree_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void inVectors_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void outVectors_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void inVectorsAbs_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void outVectorsAbs_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void ProximityAbove_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void ProximityBelow_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void ProximityLeft_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void ProximityRight_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void significanceTextBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void conditioningSetSize_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void levelsFreq_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.ConfigChanged();
        }

        private void spectralRanking_Click(object sender, RibbonControlEventArgs e)
        {
            if (this.spectralRanking.Checked)
            {
                this.allCellsFreq.Checked = false;
                this.rowCellsFreq.Checked = false;
                this.columnCellsFreq.Checked = false;
                this.levelsFreq.Checked = false;
                this.sheetFreq.Checked = true;
                this.showFixes.Enabled = true;
            } else
            {
                this.allCellsFreq.Checked = true;
                this.rowCellsFreq.Checked = true;
                this.columnCellsFreq.Checked = true;
                this.levelsFreq.Checked = false;
                this.sheetFreq.Checked = false;
            }

            setUIState(this.currentWorkbook);

            currentWorkbook.ConfigChanged();
        }

        private void spectralPlot_Click(object sender, RibbonControlEventArgs e)
        {
            // check for debug checkbox
            currentWorkbook.DebugMode = this.DebugOutput.Checked;

            // get significance threshold
            var sig = getPercent(this.significanceTextBox.Text, this.significanceTextBox.Label);

            // workbook- and UI-update callback
            Action<WorkbookState> updateWorkbook = (WorkbookState wbs) =>
            {
                this.currentWorkbook = wbs;
                setUIState(currentWorkbook);
            };

            // create progbar
            var pb = new ProgBar();

            // display form, running analysis if necessary
            currentWorkbook.showSpectralPlot(WorkbookState.MAX_DURATION_IN_MS, getConfig(), this.forceBuildDAG.Checked, pb);

            // close progbar
            pb.Close();
        }

        private void scatter3D_Click(object sender, RibbonControlEventArgs e)
        {
            // check for debug checkbox
            currentWorkbook.DebugMode = this.DebugOutput.Checked;

            // get significance threshold
            var sig = getPercent(this.significanceTextBox.Text, this.significanceTextBox.Label);

            // workbook- and UI-update callback
            Action<WorkbookState> updateWorkbook = (WorkbookState wbs) =>
            {
                this.currentWorkbook = wbs;
                setUIState(currentWorkbook);
            };

            // create progbar
            var pb = new ProgBar();

            // display form, running analysis if necessary
            currentWorkbook.show3DScatterPlot(WorkbookState.MAX_DURATION_IN_MS, getConfig(), this.forceBuildDAG.Checked, pb);

            // close progbar
            pb.Close();
        }

        private string normalizeFileName(string fileName)
        {
            // just alphanumerics
            var r = new System.Text.RegularExpressions.Regex(
                        "[^a-zA-Z0-9-_. ]",
                        System.Text.RegularExpressions.RegexOptions.Compiled
                    );

            return r.Replace(fileName, "");
        }

        private void annotate_Click(object sender, RibbonControlEventArgs e)
        {
            // if we are not currently in annotation mode:
            if (!AnnotationMode)
            {
                // get initial directory from user settings
                string iDir = (string)Properties.Settings.Default["ExceLintGroundTruthPath"];
                if (String.IsNullOrWhiteSpace(iDir))
                {
                    // default to My Documents
                    iDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                }

                // file open dialog
                var sfd = new System.Windows.Forms.SaveFileDialog();
                sfd.OverwritePrompt = false;
                sfd.DefaultExt = "csv";
                sfd.FileName = DEFAULT_GROUND_TRUTH_FILENAME;
                sfd.InitialDirectory = System.IO.Path.GetFullPath(iDir);
                sfd.RestoreDirectory = true;

                var result = sfd.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    var fileName = sfd.FileName;

                    // update user settings
                    Properties.Settings.Default["ExceLintGroundTruthPath"] = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(fileName));
                    Properties.Settings.Default.Save();

                    if (System.IO.File.Exists(fileName))
                    {
                        // append
                        Annotations = ExceLintGroundTruth.Load(fileName);
                    }
                    else
                    {
                        // otherwise, create the file
                        Annotations = ExceLintGroundTruth.Create(fileName);
                    }

                    // put notes into workbook
                    foreach (var wbs in wbstates.Values)
                    {
                        PopulateAnnotations(wbs);
                    }

                    // set the button as "stop" annotation
                    setUIState(currentWorkbook);
                }
            }
            else
            {
                // write data to file
                Annotations.Write();

                // de-populate notes
                foreach (var wbs in wbstates.Values)
                {
                    DepopulateAnnotations(wbs);
                }

                // nullify the reference
                Annotations = null;

                // set the button to "start" annotation
                setUIState(currentWorkbook);
            }
        }

        public void PopulateAnnotations(WorkbookState workbook)
        {
            if (Annotations != null)
            {
                // populate notes
                var annots = Annotations.AnnotationsFor(workbook.WorkbookName);
                foreach (var annot in annots)
                {
                    var rng = ParcelCOMShim.Address.GetCOMObject(annot.Item1, Globals.ThisAddIn.Application);
                    if (rng.Comment != null)
                    {
                        rng.Comment.Delete();
                    }
                    var comment = annot.Item2.Comment;
                    try
                    {
                        rng.AddComment(comment);
                    } catch (Exception e)
                    {
                        // for reasons unbeknownst to me, adding a comment
                        // after previously having stopped and restarted
                        // annotations sometimes throws an AccessViolationException.
                    }
                    
                }
            }
        }

        public void DepopulateAnnotations(WorkbookState workbook)
        {
            if (Annotations != null)
            {
                try
                {
                    var annots = Annotations.AnnotationsFor(workbook.WorkbookName);
                    foreach (var annot in annots)
                    {
                        var rng = ParcelCOMShim.Address.GetCOMObject(annot.Item1, Globals.ThisAddIn.Application);
                        if (rng.Comment != null && rng.Comment.Text() == annot.Item2.Comment)
                        {
                            rng.Comment.Delete();
                        }
                    }
                } catch
                {
                    // give up; this can happen as a side-effect of our wbstate removal code (cancellation timeout check)
                }
            }
            
        }

        private void annotateThisCell_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;

            // get cursor location
            var cursor = (Excel.Range)app.Selection;

            // get range object
            var rng = ParcelCOMShim.Range.RangeFromCOMObject(cursor, app.ActiveWorkbook);

            // prompt user for annotations and save results
            annotateCells(rng.Addresses(), app);
        }

        private void annotateCells(AST.Address[] addrs, Application app)
        {
            // get bug annotation from database
            var annotations = addrs.Select(addr => Annotations.AnnotationFor(addr)).ToArray();

            // populate form and ask user for new data
            var mabf = new MarkAsBugForm(annotations, addrs);

            // show form
            var result = mabf.ShowDialog();

            // pull response from form, update database and workbook
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                for (int i = 0; i < addrs.Length; i++)
                {
                    // update annotations
                    annotations[i].BugKind = mabf.BugKind;
                    annotations[i].Note = mabf.Notes;
                    Annotations.SetAnnotationFor(addrs[i], annotations[i]);

                    // get "cursor"
                    var cursor = ParcelCOMShim.Address.GetCOMObject(addrs[i], app);

                    // stick note into workbook
                    if (cursor.Comment != null)
                    {
                        cursor.Comment.Delete();
                    }
                    cursor.AddComment(annotations[i].Comment);
                }
            }
        }

        private void ExportSquareMatrixButton_Click(object sender, RibbonControlEventArgs e)
        {
            var csv = CurrentWorkbook.GetSquareMatrices(forceBuildDAG.Checked, normRefCheckBox.Checked, normSSCheckBox.Checked);
            System.Windows.Forms.Clipboard.SetText(csv);
            System.Windows.Forms.MessageBox.Show("Exported to clipboard.");
        }

        private void readClusterDump_Click(object sender, RibbonControlEventArgs e)
        {
            var ofd = new System.Windows.Forms.OpenFileDialog();
            ofd.ShowDialog();

            // read clustering from disk
            var clustering = Clustering.readClustering(ofd.FileName);

            // display
            currentWorkbook.DrawClusters(clustering);
        }

        private void moranForSelectedCells_Click(object sender, RibbonControlEventArgs e)
        {
            // get workbook
            var w = (Excel.Workbook)((Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Parent;

            // get cursor location
            var cursor = (Excel.Range)Globals.ThisAddIn.Application.Selection;

            // compute I
            var I = currentWorkbook.MoranForSelection(cursor, w, getConfig(), this.forceBuildDAG.Checked);

            // display
            System.Windows.Forms.MessageBox.Show(I.ToString());
        }

        #endregion BUTTON_HANDLERS

        #region EVENTS
        private void SetUIStateNoWorkbooks()
        {
            this.MarkAsOKButton.Enabled = false;
            this.StartOverButton.Enabled = false;
            this.AnalyzeButton.Enabled = false;
        }

        private void ExceLintRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // Callbacks for handling workbook state objects
            //WorkbookOpen(Globals.ThisAddIn.Application.ActiveWorkbook);
            //((Excel.AppEvents_Event)Globals.ThisAddIn.Application).NewWorkbook += WorkbookOpen;
            Globals.ThisAddIn.Application.WorkbookOpen += WorkbookOpen;
            Globals.ThisAddIn.Application.WorkbookActivate += WorkbookActivated;
            Globals.ThisAddIn.Application.WorkbookDeactivate += WorkbookDeactivated;
            Globals.ThisAddIn.Application.WorkbookBeforeClose += WorkbookBeforeClose;
            Globals.ThisAddIn.Application.SheetChange += SheetChange;
            Globals.ThisAddIn.Application.WorkbookAfterSave += WorkbookAfterSave;
            Globals.ThisAddIn.Application.ProtectedViewWindowOpen += ProtectedViewWindowOpen;
            Globals.ThisAddIn.Application.SheetActivate += WorksheetActivate;
            Globals.ThisAddIn.Application.SheetDeactivate += WorksheetDeactivate;
            Globals.ThisAddIn.Application.SheetSelectionChange += SheetSelectionChange;

            // sometimes the default blank workbook opens *before* the ExceLint
            // add-in is loaded so we have to handle sheet state specially.
            if (currentWorkbook == null)
            {
                var wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (wb == null)
                {
                    // the plugin loaded first; there's no active workbook
                    return;
                }
                WorkbookOpen(wb);
                WorkbookActivated(wb);
            }
        }

        private void WorksheetDeactivate(object Sh)
        {
            if (annotations != null)
            {
                annotations.Write();
            }
        }

        private void SheetSelectionChange(object Sh, Excel.Range Target)
        {
            var app = Globals.ThisAddIn.Application;

            // get cursor location
            var cursor = (Excel.Range)app.Selection;

            if (cursor.Count == 1)
            {
                // user selected a single cell
                annotateThisCell.Enabled = true;
                annotateThisCell.Label = "Annotate This Cell";
            } else if (cursor.Count > 1)
            {
                // user selected a single cell
                annotateThisCell.Enabled = true;
                annotateThisCell.Label = "Annotate These Cells";
            } else
            {
                annotateThisCell.Label = "Annotate This Cell";
                annotateThisCell.Enabled = false;
            }
        }

        private void WorksheetActivate(object Sh)
        {
            setUIState(currentWorkbook);
        }

        private void ProtectedViewWindowOpen(Excel.ProtectedViewWindow Pvw)
        {
            // set UI as nonfunctional
            setUIState(null);
        }

        private void WorkbookAfterSave(Excel.Workbook Wb, bool Success)
        {
            // this checks whether:
            // 1. there is a DAG
            // 2. the DAG changed bit is not set
            // 3. the force-update bit is not set
            if (currentWorkbook.DAGRefreshNeeded(forceDAGBuild: forceBuildDAG.Checked))
            {
                // Did the workbook really change? Diff it first.
                if (currentWorkbook.DAGChanged())
                {
                    currentWorkbook.SerializeDAG(forceDAGBuild: forceBuildDAG.Checked);
                }
            }
        }

        // This event is called when Excel opens a workbook
        private void WorkbookOpen(Excel.Workbook workbook)
        {
            WorkbookOpenHelper(workbook);
        }

        private WorkbookState WorkbookOpenHelper(Excel.Workbook workbook)
        {
            var wbs = new WorkbookState(Globals.ThisAddIn.Application, workbook);
            wbstates.Add(workbook, wbs);
            wbShutdown.AddOrUpdate(workbook, false, (k, v) => v);
            return wbs;
        }

        // This event is called when Excel brings an opened workbook
        // to the foreground
        private void WorkbookActivated(Excel.Workbook workbook)
        {
            // when opening a blank sheet, Excel does not emit
            // a WorkbookOpen event, so we need to call it manually
            if (!wbstates.ContainsKey(workbook))
            {
                WorkbookOpen(workbook);
            }
            currentWorkbook = wbstates[workbook];
            setUIState(currentWorkbook);
            if (AnnotationMode) PopulateAnnotations(currentWorkbook);
        }

        private void cancelRemoveState()
        {
            Thread.Sleep(30000);

            foreach(KeyValuePair<Excel.Workbook,Boolean> kvp in wbShutdown)
            {
                if (kvp.Value && kvp.Key != null)
                {
                    wbShutdown[kvp.Key] = false;
                }
            }
        }

        // This event is called when Excel sends an opened workbook
        // to the background
        private void WorkbookDeactivated(Excel.Workbook workbook)
        {
            if (annotations != null)
            {
                annotations.Write();
            }

            // if we recorded a workbook close for this workbook,
            // remove all workbook state
            if (wbShutdown[workbook])
            {
                wbstates.Remove(workbook);
                if (wbstates.Count == 0)
                {
                    SetUIStateNoWorkbooks();
                }
                Boolean outcome;
                wbShutdown.TryRemove(workbook, out outcome);
            }

            if (AnnotationMode) DepopulateAnnotations(currentWorkbook);

            currentWorkbook = null;

            // WorkbookBeforeClose event does not fire for default workbooks
            // containing no data
            var wbs = new List<Excel.Workbook>();
            foreach (var wb in Globals.ThisAddIn.Application.Workbooks)
            {
                if (wb != workbook)
                {
                    wbs.Add((Excel.Workbook)wb);
                }
            }

            if (wbs.Count == 0)
            {
                wbstates.Clear();
                SetUIStateNoWorkbooks();
            }
        }

        private void WorkbookBeforeClose(Excel.Workbook workbook, ref bool Cancel)
        {
            // record possible workbook close
            wbShutdown[workbook] = true;
            // if the user hits cancel, then WorkbookDeactivated will never
            // be called, but the shutdown will still be recorded; the
            // following schedules an action to unrecord the workbook close
            Thread t = new Thread(cancelRemoveState);
            t.Start();
        }

        private void SheetChange(object worksheet, Excel.Range target)
        {
            if (currentWorkbook != null)
            {
                currentWorkbook.MarkDAGAsChanged();
                currentWorkbook.resetTool();
                setUIState(currentWorkbook);
            }
        }
        #endregion EVENTS

        #region UTILITY_FUNCTIONS
        private static FSharpOption<double> getPercent(string input, string label)
        {
            var errormsg = label + " must be a value between 0 and 100.";

            try
            {
                double prop = Double.Parse(input) / 100.0;

                if (prop <= 0 || prop > 100)
                {
                    System.Windows.Forms.MessageBox.Show(errormsg);
                }

                return FSharpOption<double>.Some(prop);
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show(errormsg);
                return FSharpOption<double>.None;
            }
        }

        private void SetTooltips(string text)
        {
            this.MarkAsOKButton.ScreenTip = text;
            this.StartOverButton.ScreenTip = text;
            this.AnalyzeButton.ScreenTip = text;
            this.showHeatmap.ScreenTip = text;
            this.allCellsFreq.ScreenTip = text;
            this.columnCellsFreq.ScreenTip = text;
            this.rowCellsFreq.ScreenTip = text;
            this.levelsFreq.ScreenTip = text;
            this.sheetFreq.ScreenTip = text;
            this.DebugOutput.ScreenTip = text;
            this.forceBuildDAG.ScreenTip = text;
            this.inferAddrModes.ScreenTip = text;
            this.allCells.ScreenTip = text;
            this.weightByIntrinsicAnomalousness.ScreenTip = text;
            this.significanceTextBox.ScreenTip = text;
            this.conditioningSetSize.ScreenTip = text;
            this.spectralRanking.ScreenTip = text;
        }

        private void setUIState(WorkbookState wbs)
        {
            var isAChart = false;
            var sheetProtected = false;
            Worksheet w = null;

            try
            {
                w = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                sheetProtected = wbs.WorksheetProtected(w);
            } catch
            {
                isAChart = true;
            }

            if (wbs == null || Globals.ThisAddIn.Application.ActiveProtectedViewWindow != null || sheetProtected || isAChart)
            {
                // disable all controls
                var disabled = false;
                var disabled_text = "ExceLint is disabled in protected mode.  Please enable editing to continue.";

                this.MarkAsOKButton.Enabled = disabled;
                this.StartOverButton.Enabled = disabled;
                this.AnalyzeButton.Enabled = disabled;
                this.showHeatmap.Enabled = disabled;
                this.allCellsFreq.Enabled = disabled;
                this.columnCellsFreq.Enabled = disabled;
                this.rowCellsFreq.Enabled = disabled;
                this.levelsFreq.Enabled = disabled;
                this.DebugOutput.Enabled = disabled;
                this.forceBuildDAG.Enabled = disabled;
                this.inferAddrModes.Enabled = disabled;
                this.allCells.Enabled = disabled;
                this.sheetFreq.Enabled = disabled;
                this.weightByIntrinsicAnomalousness.Enabled = disabled;
                this.significanceTextBox.Enabled = disabled;
                this.conditioningSetSize.Enabled = disabled;
                this.spectralRanking.Enabled = disabled;
                this.showFixes.Enabled = disabled;
                this.annotate.Enabled = disabled;
                this.useResultant.Enabled = disabled;
                this.ClusterBox.Enabled = disabled;
                this.RunCUSTODES.Enabled = disabled;

                // tell the user ExceLint doesn't work
                SetTooltips(disabled_text);
            } else
            {
                // clear button text
                SetTooltips("");

                // enable auditing buttons if an audit has started
                this.MarkAsOKButton.Enabled = wbs.MarkAsOK_Enabled;
                this.StartOverButton.Enabled = wbs.ClearColoringButton_Enabled;
                this.AnalyzeButton.Enabled = wbs.Analyze_Enabled && wbs.Visualization_Hidden(w);

                // only enable viewing heatmaps if we are not in the middle of an analysis
                this.showHeatmap.Enabled = wbs.Analyze_Enabled && wbs.CUSTODES_Hidden(w);
                this.RunCUSTODES.Enabled = wbs.Analyze_Enabled && wbs.Visualization_Hidden(w);

                // disable config buttons if we are:
                // 1. in the middle of an audit, or
                // 2. we are viewing the heatmap, or
                // 3. if spectral ranking is checked, disable scopes
                var enable_config = wbs.Analyze_Enabled && wbs.Visualization_Hidden(w) && wbs.CUSTODES_Hidden(w);
                this.allCellsFreq.Enabled = enable_config && !this.spectralRanking.Checked;
                this.columnCellsFreq.Enabled = enable_config && !this.spectralRanking.Checked;
                this.rowCellsFreq.Enabled = enable_config && !this.spectralRanking.Checked;
                this.levelsFreq.Enabled = enable_config && !this.spectralRanking.Checked;
                this.sheetFreq.Enabled = enable_config && !this.spectralRanking.Checked;
                this.DebugOutput.Enabled = enable_config;
                this.forceBuildDAG.Enabled = enable_config;
                this.inferAddrModes.Enabled = enable_config;
                this.allCells.Enabled = enable_config;
                this.weightByIntrinsicAnomalousness.Enabled = enable_config;
                this.significanceTextBox.Enabled = enable_config;
                this.conditioningSetSize.Enabled = enable_config;
                this.spectralRanking.Enabled = enable_config;
                this.useResultant.Enabled = enable_config;
                this.showFixes.Enabled = enable_config && this.spectralRanking.Checked;
                this.annotate.Enabled = true;   // user can annotate at any time
                this.ClusterBox.Enabled = enable_config;

                // toggle the heatmap label depending on the heatmap shown/hidden state
                if (wbs.Visualization_Hidden(w))
                {
                    this.showHeatmap.Label = "Show Formula Similarity";
                }
                else
                {
                    this.showHeatmap.Label = "Hide Formula Similarity";
                }

                // toggle the CUSTODES label depending on the analysis shown/hidden state
                if (wbs.CUSTODES_Hidden(w))
                {
                    this.RunCUSTODES.Label = "Run CUSTODES";
                }
                else
                {
                    this.RunCUSTODES.Label = "Hide CUSTODES";
                }

                // toggle the annotation button depending on whether we have
                // an annotation datastructure open or not
                if (AnnotationMode)
                {
                    this.annotate.Label = "Stop Annotating";
                    this.annotateThisCell.Visible = true;
                } else
                {
                    this.annotate.Label = "Annotate";
                    this.annotateThisCell.Visible = false;
                }
            }
        }

        public ExceLint.FeatureConf getConfig()
        {
            var c = new ExceLint.FeatureConf();

            // spatiostructual vectors
            if (this.ClusterBox.Checked)
            {
                c = c.enableShallowInputVectorMixedFullCVectorResultantOSI(true);
            } else if (this.useResultant.Checked) {
                c = c.enableShallowInputVectorMixedResultant(true);
            } else {
                c = c.enableShallowInputVectorMixedL2NormSum(true);
            }

            // Scopes (i.e., conditioned analysis)
            if (this.allCellsFreq.Checked) { c = c.analyzeRelativeToAllCells(true); }
            if (this.columnCellsFreq.Checked) { c = c.analyzeRelativeToColumns(true); }
            if (this.rowCellsFreq.Checked) { c = c.analyzeRelativeToRows(true); }
            if (this.levelsFreq.Checked) { c = c.analyzeRelativeToLevels(true); }
            if (this.sheetFreq.Checked) { c = c.analyzeRelativeToSheet(true); }

            // weighting / program resynthesis
            if (this.inferAddrModes.Checked) { c = c.inferAddressModes(true);  }
            if (!this.allCells.Checked) { c = c.analyzeOnlyFormulas(true);  }
            if (this.weightByIntrinsicAnomalousness.Checked) { c = c.weightByIntrinsicAnomalousness(true); }
            if (this.conditioningSetSize.Checked) { c = c.weightByConditioningSetSize(true); }

            // ranking type
            if (this.spectralRanking.Checked) { c = c.spectralRanking(true).analyzeRelativeToSheet(true); }

            // distance metric
            switch(distanceCombo.Text)
            {
                case "Earth Mover":
                    c = c.enableDistanceEarthMover(true);
                    break;
                case "Nearest Neighbor":
                    c = c.enableDistanceNearestNeighbor(true);
                    break;
                case "Mean Centroid":
                    c = c.enableDistanceMeanCentroid(true);
                    break;
            }   
            
            // COF?
            c = c.enableShallowInputVectorMixedCOFRefUnnormSSNorm(false);

            // limit analysis to a single sheet
            c = c.limitAnalysisToSheet(((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet).Name);

            // debug mode
            if (this.DebugOutput.Checked) { c = c.enableDebugMode(true); }

            return c;
        }

        public bool AnnotationMode
        {
            get { return annotations != null; }
        }

        public ExceLintGroundTruth Annotations
        {
            get { return annotations; }
            set { annotations = value; }
        }

        #endregion UTILITY_FUNCTIONS
    }
}

