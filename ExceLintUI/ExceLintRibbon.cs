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

        private void RegularityMap_Click(object sender, RibbonControlEventArgs e)
        {
            // disable annoying OLE warnings
            Globals.ThisAddIn.Application.DisplayAlerts = false;

            // get dependence graph
            var graph = currentWorkbook.getDependenceGraph(false);

            // get active sheet
            Worksheet activeWs = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            if (fixClusterModel == null)
            {
                // change button name
                RegularityMap.Label = "Hide Global View";

                // create progbar in main thread;
                // worker thread will call Dispose
                var pb = new ProgBar();

                // build the model
                var sw2 = System.Diagnostics.Stopwatch.StartNew();
                fixClusterModel = currentWorkbook.NewEntropyModelForWorksheet2(activeWs, getConfig(), true, pb);

                // get z for worksheet
                int z = -1;
                try
                {
                    z = fixClusterModel.ZForWorksheet(activeWs.Name);
                } catch (KeyNotFoundException)
                {
                    pb.Close();
                    return;
                }

                // do visualization
                var histo2 = fixClusterModel.InvertedHistogram;
                var clusters2 = fixClusterModel.Clustering(z);
                sw2.Stop();
                var cl_filt2 = PrettyClusters(clusters2, histo2, graph);

                // in case we've run something before, restore colors
                currentWorkbook.restoreOutputColors();

                // save colors
                CurrentWorkbook.saveColors(activeWs);

                // paint
                currentWorkbook.DrawImmutableClusters(cl_filt2, fixClusterModel.InvertedHistogram, activeWs);

                // remove progress bar
                pb.Close();
            } else
            {
                currentWorkbook.restoreOutputColors();

                // change button name
                RegularityMap.Label = "Show Global View";

                // reset model
                fixClusterModel = null;
            }
        }

        private void ClearEverything_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.resetTool();
            setUIState(currentWorkbook);
        }
        
        private void ExceLintVsTrueSmells_Click(object sender, RibbonControlEventArgs e)
        {
            try
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

                    // get excelint clusterings
                    Worksheet activeWs = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                    var model = ModelInit(activeWs);
                    var eclusters = GetEntropyClustering(model, activeWs);

                    // get clustering diff
                    var diff = clusteringDiff(eclusters, clustering);

                    Globals.ThisAddIn.Application.ScreenUpdating = false;

                    // restore colors
                    CurrentWorkbook.restoreOutputColors();

                    // save colors
                    CurrentWorkbook.saveColors(activeWs);

                    // clear colors
                    CurrentWorkbook.ClearAllColors(activeWs);

                    // draw all the ExceLint flags blue
                    foreach (var addr in diff.Item1)
                    {
                        currentWorkbook.paintColor(addr, System.Drawing.Color.Blue);
                    }

                    // draw all the intersecting cells purple
                    foreach (var addr in diff.Item2)
                    {
                        currentWorkbook.paintColor(addr, System.Drawing.Color.Purple);
                    }

                    // draw all the true smells blredue
                    foreach (var addr in diff.Item3)
                    {
                        currentWorkbook.paintColor(addr, System.Drawing.Color.Red);
                    }

                    Globals.ThisAddIn.Application.ScreenUpdating = true;

                    System.Windows.Forms.MessageBox.Show("excelint-only: blue\nboth: purple\ntrue smell: red");
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Can't find true smells CSV or workbook directory");
                }
            } catch (Exception)
            {
                // any exception; clear properties and start over
                Properties.Settings.Default.Reset();
                ExceLintVsTrueSmells_Click(sender, e);
            }
        }

        private Tuple<HashSet<AST.Address>, HashSet<AST.Address>, HashSet<AST.Address>> clusteringDiff(IEnumerable<IEnumerable<AST.Address>> eclusters, IEnumerable<IEnumerable<AST.Address>> otherclusters)
        {
            // flatten both
            var ecells = new HashSet<AST.Address>(eclusters.SelectMany(i => i));
            var ocells = new HashSet<AST.Address>(otherclusters.SelectMany(i => i));

            // find the set in e but not in o
            var eonly = new HashSet<AST.Address>(ecells.Except(ocells));

            // find the set in o but not in e
            var oonly = new HashSet<AST.Address>(ocells.Except(ecells));

            // find the set in both
            var both = new HashSet<AST.Address>(ecells.Intersect(ocells));

            return new Tuple<HashSet<AST.Address>, HashSet<AST.Address>, HashSet<AST.Address>>(eonly, both, oonly);
        }

        private string ProposedFixesToString(CommonTypes.ProposedFix[] fixes)
        {
            // produce output string
            var sb = new StringBuilder();

            sb.Append("SOURCE");
            sb.Append(" -> ");
            sb.Append("TARGET");
            sb.Append(" = ");
            sb.Append("- TARGET DISTANCE");
            sb.Append(" / ");
            sb.Append("(");
            sb.Append(" ENTROPY_DELTA ");
            sb.Append(" * ");
            sb.Append(" DISTANCE ");
            sb.Append(")");
            sb.Append(" = ");
            sb.Append(" RESULT ");
            sb.Append(", FIX FREQ ");
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
                sb.Append("-");
                sb.Append(fix.TargetSize.ToString());
                sb.Append(" / ");
                sb.Append("(");
                sb.Append(fix.EntropyDelta.ToString());
                sb.Append(" * ");
                sb.Append(fix.Distance.ToString());
                sb.Append(")");                
                sb.Append(" = ");
                sb.Append(fix.Score.ToString());

                // fix freq
                sb.Append(" , ");
                sb.Append(fix.FixFrequency.ToString());
                sb.Append(" ,");
                sb.Append(fix.FixFrequencyScore.ToString());

                // EOL
                sb.AppendLine();
            }

            return sb.ToString();
        }

        private Clusters GetEntropyClustering(EntropyModelBuilder2.EntropyModel2 model, Worksheet w)
        {
            // get z for this worksheet
            var z = model.ZForWorksheet(w.Name);

            // get fixes
            var fixes = model.Fixes(z);

            // get ranking
            var ranking = EntropyModelBuilder2.EntropyModel2.Ranking(fixes);

            // extract clusters
            var clusters = EntropyModelBuilder2.EntropyModel2.RankingToClusters(fixes);

            return clusters;
        }

        private EntropyModelBuilder2.EntropyModel2 ModelInit(Worksheet activeWs)
        {
            // get config
            var conf = getConfig();

            // get significance threshold
            conf = conf.setThresh(0.95);

            // create progbar in main thread;
            // worker thread will call Dispose
            var pb = new ProgBar();

            // build the model
            var model = currentWorkbook.NewEntropyModelForWorksheet2(activeWs, conf, true, pb);

            // remove progress bar
            pb.Close();

            return model;
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
            var cs1 = ElideStringClusters(cs, ih, graph);
            var cs2 = ElideWhitespaceClusters(cs1, ih, graph);
            return cs2;
        }

        private void VectorForCell_Click(object sender, RibbonControlEventArgs e)
        {
            // get cursor location
            var cursor = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            AST.Address cursorAddr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, Globals.ThisAddIn.Application.ActiveWorkbook);

            // get config
            var conf = getConfig();

            // get dependence graph
            var dag = currentWorkbook.getDependenceGraph(true);

            var sb = new StringBuilder();

            // get vector for each enabled feature
            var feats = conf.EnabledFeatures;
            for (int i = 0; i < feats.Length; i++)
            {
                // run feature
                //sb.Append(feats[i]);
                //sb.Append(" = ");
                sb.Append(conf.get_FeatureByName(feats[i]).Invoke(cursorAddr).Invoke(dag).ToString());
                //sb.Append("\n");
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
            // test
            var graph = new FastDependenceAnalysis.Graph(Globals.ThisAddIn.Application, (Worksheet)Globals.ThisAddIn.Application.ActiveSheet);

            // check for debug checkbox
            currentWorkbook.DebugMode = false;

            // workbook- and UI-update callback
            Action<WorkbookState> updateWorkbook = (WorkbookState wbs) =>
            {
                this.currentWorkbook = wbs;
                setUIState(currentWorkbook);
            };

            // create progbar in main thread;
            // worker thread will call Dispose
            var pb = new ProgBar();

            Worksheet activeWs = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            DoAnalysis(0.95, currentWorkbook, getConfig(), true, updateWorkbook, pb, true, activeWs);
        }

        public static void DoAnalysis(FSharpOption<double> sigThresh, WorkbookState wbs, FeatureConf conf, bool forceBuildDAG, Action<WorkbookState> updateState, ProgBar pb, bool showFixes, Worksheet ws)
        {
            if (sigThresh == FSharpOption<double>.None)
            {
                return;
            }
            else
            {
                wbs.toolSignificance = sigThresh.Value;
                conf = conf.setThresh(sigThresh.Value);
                try
                {
                    wbs.analyze(WorkbookState.MAX_DURATION_IN_MS, conf, forceBuildDAG, pb);
                    wbs.MarkAsOK_Enabled = true;
                    wbs.setTool(active: true);
                    wbs.saveColors(ws);
                    wbs.flagNext(ws);
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
            Worksheet activeWs = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            currentWorkbook.markAsOK(true, activeWs);
            setUIState(currentWorkbook);
        }

        private void StartOverButton_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.resetTool();
            setUIState(currentWorkbook);
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
            if (currentWorkbook.DAGRefreshNeeded(forceDAGBuild: true))
            {
                // Did the workbook really change? Diff it first.
                if (currentWorkbook.DAGChanged())
                {
                    currentWorkbook.SerializeDAG(forceDAGBuild: true);
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
            WorkbookState wbs;
            if (!wbstates.ContainsKey(workbook))
            {
                wbs = new WorkbookState(Globals.ThisAddIn.Application, workbook);
                wbstates.Add(workbook, wbs);
                wbShutdown.AddOrUpdate(workbook, false, (k, v) => v);
            } else
            {
                wbs = wbstates[workbook];
            }
            
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
                this.RegularityMap.Enabled = wbs.Analyze_Enabled && wbs.Visualization_Hidden(w);

                // disable config buttons if we are:
                // 1. in the middle of an audit, or
                // 2. we are viewing the heatmap, or
                // 3. if spectral ranking is checked, disable scopes
                var enable_config = wbs.Analyze_Enabled && wbs.Visualization_Hidden(w) && wbs.CUSTODES_Hidden(w);
            }
        }

        public ExceLint.FeatureConf getConfig()
        {
            var c = new ExceLint.FeatureConf();


            c = c.enableShallowInputVectorMixedFullCVectorResultantOSI(true);
            
            // limit analysis to a single sheet
            c = c.limitAnalysisToSheet(((Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet).Name);

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

