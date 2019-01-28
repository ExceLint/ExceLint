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
using Graph = FastDependenceAnalysis.Graph;

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

        #region BUTTON_HANDLERS

        private void RegularityMap_Click(object sender, RibbonControlEventArgs e)
        {
            // disable annoying OLE warnings
            Globals.ThisAddIn.Application.DisplayAlerts = false;

            // get dependence graph
            var graph = new Graph(Globals.ThisAddIn.Application, (Worksheet)Globals.ThisAddIn.Application.ActiveSheet);

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
                fixClusterModel = currentWorkbook.NewEntropyModelForWorksheet2(activeWs, getConfig(), graph, pb);

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
       
        private Clusters ElideWhitespaceClusters(Clusters cs, ROInvertedHistogram ih, Graph graph)
        {
            var output = new HashSet<ImmutableHashSet<AST.Address>>();

            foreach (ImmutableHashSet<AST.Address> c in cs)
            {
                if (!c.All(a => EntropyModelBuilder2.AddressIsWhitespaceValued(a, ih, graph)))
                {
                    output.Add(c);
                }
            }

            return output.ToImmutableHashSet();
        }

        private Clusters ElideStringClusters(Clusters cs, ROInvertedHistogram ih, Graph graph)
        {
            var output = new HashSet<ImmutableHashSet<AST.Address>>();

            foreach (ImmutableHashSet<AST.Address> c in cs)
            {
                if (!c.All(a => EntropyModelBuilder2.AddressIsStringValued(a, ih, graph)))
                {
                    output.Add(c);
                }
            }

            return output.ToImmutableHashSet();
        }

        private Clusters PrettyClusters(Clusters cs, ROInvertedHistogram ih, Graph graph)
        {
            var cs1 = ElideStringClusters(cs, ih, graph);
            var cs2 = ElideWhitespaceClusters(cs1, ih, graph);
            return cs2;
        }

        public WorkbookState CurrentWorkbook { 
            get
            {
                return currentWorkbook;
            }
        }

        private void AnalyzeButton_Click(object sender, RibbonControlEventArgs e)
        {
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
            Globals.ThisAddIn.Application.WorkbookOpen += WorkbookOpen;
            Globals.ThisAddIn.Application.WorkbookActivate += WorkbookActivated;
            Globals.ThisAddIn.Application.WorkbookDeactivate += WorkbookDeactivated;
            Globals.ThisAddIn.Application.WorkbookBeforeClose += WorkbookBeforeClose;
            Globals.ThisAddIn.Application.SheetChange += SheetChange;
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
            //var app = Globals.ThisAddIn.Application;

            //// get cursor location
            //var cursor = (Excel.Range)app.Selection;
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
                currentWorkbook.resetTool();
                setUIState(currentWorkbook);
            }
        }
        #endregion EVENTS

        #region UTILITY_FUNCTIONS
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

        public FeatureConf getConfig()
        {
            var c = new FeatureConf();
            c = c.enableShallowInputVectorMixedFullCVectorResultantOSI(true);
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

