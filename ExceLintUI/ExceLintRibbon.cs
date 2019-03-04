using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Threading;
using ExceLintFileFormats;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using ExceLint;
using System.Collections.Immutable;
using FastDependenceAnalysis;
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
            if (currentWorkbook.Analyze_Enabled)
            {
                // disable annoying OLE warnings
                Globals.ThisAddIn.Application.DisplayAlerts = false;

                // get dependence graph
                var graph = new Graph(Globals.ThisAddIn.Application, (Worksheet)Globals.ThisAddIn.Application.ActiveSheet);

                // get active sheet
                Worksheet activeWs = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

                // get cells
                var cells = graph.allCells();

                // config
                var conf = getConfig();

                // get fingerprints
                // I am hard-coding the one feature here for now, since
                // I have removed all of the other ExceLint features
                var ns = CommonFunctions.runEnabledFeatures(cells, graph, conf, Progress.NOPProgress());
                var fs = ns["ShallowInputVectorMixedFullCVectorResultantOSI"];

                // cluster
                var cs = ClusterFingerprints(fs);

                // get inverted histogram
                var histo = CommonFunctions.invertedHistogram(ns, graph, conf);


                // filter out whitespace and string clusters
                var cs_filtered = PrettyClusters(cs, histo, graph);

                // in case we've run something before, restore colors
                currentWorkbook.restoreOutputColors();

                // save colors
                CurrentWorkbook.saveColors(activeWs);

                // paint formulas
                var colormap = currentWorkbook.DrawImmutableClusters(cs_filtered, histo, activeWs, graph, this.analyzeFormulas.Checked);

                // if checked, analyze data
                if (this.enableDataHighlight.Checked)
                {
                    // get formula usage for data cells
                    var referents = Vector.getReferentDict(graph);

                    // paint data
                    currentWorkbook.ColorDataWithMap(referents, colormap, graph);

                    // add data comments
                    currentWorkbook.labelReferents(referents);
                }

                // set UI state
                setUIState(currentWorkbook);
            }
            else
            {
                // clear
                currentWorkbook.restoreOutputColors();

                // set UI state
                setUIState(currentWorkbook);
            }
           
        }

        private Clusters ClusterFingerprints(Tuple<AST.Address, Countable>[] fs)
        {
            // group by fingerprint and turn into ImmutableClustering
            var cs = fs.GroupBy(kvp => kvp.Item2.LocationFree);
            var ctmp = new HashSet<ImmutableHashSet<AST.Address>>();
            foreach (var group in cs)
            {
                var hs = new HashSet<AST.Address>();
                foreach (var tup in group.AsEnumerable())
                {
                    var addr = tup.Item1;
                    hs.Add(addr);
                }

                var ihs = hs.ToImmutableHashSet();

                ctmp.Add(ihs);
            }

            return ctmp.ToImmutableHashSet();
        }

        /**
         * For troubleshooting colorstops, which are largely undocumented
         */
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;

            // get cursor location
            var cursor = (Excel.Range)app.Selection;

            if (cursor.Count == 1)
            {
                // user selected a single cell
                var gradient = (cursor.Interior.Gradient as Excel.LinearGradient);
                var cs = gradient.ColorStops;
                String s = "";
                foreach (var c in cs)
                {
                    var c1 = (Excel.ColorStop)c;
                    s += "color: " + c1.Color.ToString() + "\n";
                    s += "tintandshade: " + c1.TintAndShade.ToString() + "\n";
                    s += "position: " + c1.Position.ToString() + "\n";
                    s += "themecolor: " + c1.ThemeColor.ToString() + "\n";
                    s += "------------\n";
                }
                System.Windows.Forms.MessageBox.Show(s);
            }

        }

        private void analyzeFormulas_Click(object sender, RibbonControlEventArgs e)
        {
            setUIState(currentWorkbook);
        }

        private void enableDataHighlight_Click(object sender, RibbonControlEventArgs e)
        {
            setUIState(currentWorkbook);
        }
        #endregion BUTTON_HANDLERS

        #region EVENTS
        private void SetUIStateNoWorkbooks()
        {
            this.RegularityMap.Enabled = false;
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
            this.RegularityMap.ScreenTip = text;
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
                var disabled_text = "ExceLint is disabled in protected mode.  Please enable editing to continue.";

                // tell the user ExceLint doesn't work
                SetTooltips(disabled_text);
            } else
            {
                // clear button text
                SetTooltips("");

                // button should never be locked here
                // (gets locked when there is no workbook open)
                this.RegularityMap.Enabled = true;

                // only enable reveal structure button if at least
                // one checkbox is checked
                if (this.enableDataHighlight.Checked || this.analyzeFormulas.Checked)
                {
                    this.RegularityMap.Enabled = true;
                } else
                {
                    this.RegularityMap.Enabled = false;
                }

                // only enable viewing heatmaps if we are not in the middle of an analysis
                if (wbs.Analyze_Enabled)
                {
                    this.RegularityMap.Label = "Reveal Structure";
                    this.enableDataHighlight.Enabled = true;
                    this.analyzeFormulas.Enabled = true;
                }
                else
                {
                    this.RegularityMap.Label = "Hide Structure";
                    this.enableDataHighlight.Enabled = false;
                    this.analyzeFormulas.Enabled = false;
                }
            }
        }

        public FeatureConf getConfig()
        {
            var c = new FeatureConf();
            c = c.enableShallowInputVectorMixedFullCVectorResultantOSI(true);
            return c.validate;
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

        public WorkbookState CurrentWorkbook
        {
            get { return currentWorkbook; }
        }

        #endregion UTILITY_FUNCTIONS
    }
}

