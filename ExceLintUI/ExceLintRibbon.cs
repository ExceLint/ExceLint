using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.FSharp.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;

namespace ExceLintUI
{
    public partial class ExceLintRibbon
    {
        Dictionary<Excel.Workbook, WorkbookState> wbstates = new Dictionary<Excel.Workbook, WorkbookState>();
        System.Collections.Concurrent.ConcurrentDictionary<Excel.Workbook,Boolean> wbShutdown = new System.Collections.Concurrent.ConcurrentDictionary<Excel.Workbook,Boolean>();
        WorkbookState currentWorkbook;

        #region BUTTON_HANDLERS
        private void showVectors_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.getMixedFormulaVectors(forceDAGBuild: forceBuildDAG.Checked);
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
            DoAnalysis(sig, currentWorkbook, getConfig(), this.forceBuildDAG.Checked, updateWorkbook, pb);
        }

        private static void DoAnalysis(FSharpOption<double> sigThresh, WorkbookState wbs, ExceLint.FeatureConf conf, bool forceBuildDAG, Action<WorkbookState> updateState, ProgBar pb)
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
                    wbs.flag();
                    updateState(wbs);

                    var debug_info = prepareDebugInfo(wbs);
                    var timing_info = prepareTimingInfo(wbs);

                    // debug output
                    if (wbs.DebugMode)
                    {
                        RunInSTAThread(() => printDebugInfo(debug_info, timing_info));
                    }

                    pb.GoAway();
                }
                catch (Parcel.ParseException ex)
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

            if (FSharpOption<WorkbookState.Analysis>.get_IsNone(a))
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
                var causes = analysis.model.causeOf(score.Key);
                var causes_str = "\tcauses: [\n" + String.Join("\n", causes.Select(cause => "\t\t" + ExceLint.ErrorModel.prettyHistoBinDesc(cause.Key) + " = " + cause.Value)) + "\n\t]";

                // print
                return prefix + score.Key.A1FullyQualified() + " -> " + score.Value.ToString() + "\n" + causes_str + "\n\t" + "weight: " + analysis.model.weightOf(score.Key);
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

            if (FSharpOption<WorkbookState.Analysis>.get_IsNone(a))
            {
                return "";
            }

            var analysis = a.Value;

            // time and space information
            var time_str = "DAG construction ms: " + analysis.dag.AnalysisMilliseconds + "\n" +
                           "Feature scoring ms: " + analysis.model.ScoreTimeInMilliseconds + "\n" +
                           "Num score entries: " + analysis.model.NumScoreEntries + "\n" +
                           "Frequency counting ms: " + analysis.model.FrequencyTableTimeInMilliseconds + "\n" +
                           "Num freq table entries: " + analysis.model.NumFreqEntries + "\n" +
                           "Ranking ms: " + analysis.model.RankingTimeInMilliseconds + "\n" +
                           "Total ranking length: " + analysis.model.NumRankedEntries;

            return time_str;
        }

        private static void printDebugInfo(string debug_info, string time_info)
        {
            if (!String.IsNullOrEmpty(debug_info))
            {
                System.Windows.Forms.Clipboard.SetText(debug_info);
                System.Windows.Forms.MessageBox.Show(debug_info);
            }

            if (!String.IsNullOrEmpty(debug_info))
            {
                System.Windows.Forms.Clipboard.SetText(time_info);
                System.Windows.Forms.MessageBox.Show(time_info);
            }
        }

        private static void RunInSTAThread(ThreadStart t)
        {
            Thread thread = new Thread(t);
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        }

        private void MarkAsOKButton_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.markAsOK();
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
            currentWorkbook.getSelected(getConfig(), Scope.Selector.SameColumn, forceDAGBuild: forceBuildDAG.Checked);
        }

        private void rowSelected_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.getSelected(getConfig(), Scope.Selector.SameRow, forceDAGBuild: forceBuildDAG.Checked);
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            // if the user holds down Option, get absolute vectors
            if ((System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Alt) > 0)
            {
                currentWorkbook.getRawFormulaVectors(forceDAGBuild: forceBuildDAG.Checked);
            } else
            {
                currentWorkbook.getFormulaRelVectors(forceDAGBuild: forceBuildDAG.Checked);
            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            // if the user holds down Option, get absolute vectors
            if ((System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Alt) > 0)
            {
                currentWorkbook.getRawDataVectors(forceDAGBuild: forceBuildDAG.Checked);
            }
            else
            {
                currentWorkbook.getDataRelVectors(forceDAGBuild: forceBuildDAG.Checked);
            }
        }

        private void FrmAbsVect_Click(object sender, RibbonControlEventArgs e)
        {
            // if the user holds down Option, get absolute vectors
            if ((System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Alt) > 0)
            {
                currentWorkbook.getRawFormulaVectors(forceDAGBuild: forceBuildDAG.Checked);
            }
            else
            {
                currentWorkbook.getFormulaAbsVectors(forceDAGBuild: forceBuildDAG.Checked);
            }
        }

        private void DataAbsVect_Click(object sender, RibbonControlEventArgs e)
        {
            // if the user holds down Option, get absolute vectors
            if ((System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Alt) > 0)
            {
                currentWorkbook.getRawDataVectors(forceDAGBuild: forceBuildDAG.Checked);
            }
            else
            {
                currentWorkbook.getDataAbsVectors(forceDAGBuild: forceBuildDAG.Checked);
            }
        }

        private void showHeatmap_Click(object sender, RibbonControlEventArgs e)
        {
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
            //Task t = Task.Run(() => DoHeatmap(sig, currentWorkbook, getConfig(), this.forceBuildDAG.Checked, updateWorkbook, pb));
            DoHeatmap(sig, currentWorkbook, getConfig(), this.forceBuildDAG.Checked, updateWorkbook, pb);
        }

        private static void DoHeatmap(FSharpOption<double> sigThresh, WorkbookState wbs, ExceLint.FeatureConf conf, bool forceBuildDAG, Action<WorkbookState> updateState, ProgBar pb)
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
                    wbs.toggleHeatMap(WorkbookState.MAX_DURATION_IN_MS, conf, forceBuildDAG, pb);
                    updateState(wbs);

                    var debug_info = prepareDebugInfo(wbs);
                    var timing_info = prepareTimingInfo(wbs);

                    // debug output
                    if (wbs.DebugMode)
                    {
                        RunInSTAThread(() => printDebugInfo(debug_info, timing_info));
                    }

                    pb.GoAway();
                }
                catch (Parcel.ParseException ex)
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

            // sometimes the default blank workbook opens *before* the CheckCell
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

        private void WorkbookAfterSave(Excel.Workbook Wb, bool Success)
        {
            currentWorkbook.SerializeDAG(forceDAGBuild: forceBuildDAG.Checked);
        }

        // This event is called when Excel opens a workbook
        private void WorkbookOpen(Excel.Workbook workbook)
        {
            wbstates.Add(workbook, new WorkbookState(Globals.ThisAddIn.Application, workbook));
            wbShutdown.AddOrUpdate(workbook, false, (k, v) => v);
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
            System.Threading.Thread.Sleep(1000);

            foreach(KeyValuePair<Excel.Workbook,Boolean> kvp in wbShutdown)
            {
                if (kvp.Value)
                {
                    wbShutdown[kvp.Key] = false;
                }
            }
        }

        // This even it called when Excel sends an opened workbook
        // to the background
        private void WorkbookDeactivated(Excel.Workbook workbook)
        {
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
            System.Threading.Thread t = new System.Threading.Thread(cancelRemoveState);
            t.Start();
        }

        private void SheetChange(object worksheet, Excel.Range target)
        {
            currentWorkbook.DAGChanged();
            currentWorkbook.resetTool();
            setUIState(currentWorkbook);
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

        private void setUIState(WorkbookState wbs)
        {
            // enable auditing buttons if an audit has started
            this.MarkAsOKButton.Enabled = wbs.MarkAsOK_Enabled;
            this.StartOverButton.Enabled = wbs.ClearColoringButton_Enabled;
            this.AnalyzeButton.Enabled = wbs.Analyze_Enabled && wbs.HeatMap_Hidden;

            // only enable viewing heatmap if we are not in the middle of an analysis
            this.showHeatmap.Enabled = wbs.Analyze_Enabled;

            // disable config buttons if we are:
            // 1. in the middle of an audit, or
            // 2. we are viewing the heatmap
            var enable_config = wbs.Analyze_Enabled && wbs.HeatMap_Hidden;
            this.allCellsFreq.Enabled = enable_config;
            this.columnCellsFreq.Enabled = enable_config;
            this.rowCellsFreq.Enabled = enable_config;
            this.levelsFreq.Enabled = enable_config;
            this.DebugOutput.Enabled = enable_config;
            this.forceBuildDAG.Enabled = enable_config;
            this.inferAddrModes.Enabled = enable_config;
            this.allCells.Enabled = enable_config;
            this.weightByIntrinsicAnomalousness.Enabled = enable_config;
            this.significanceTextBox.Enabled = enable_config;

            // toggle the heatmap label depending on the heatmap shown/hidden state
            if (wbs.HeatMap_Hidden)
            {
                this.showHeatmap.Label = "Show Heat Map";
            } else
            {
                this.showHeatmap.Label = "Hide Heat Map";
            }
        }

        public ExceLint.FeatureConf getConfig()
        {
            var c = new ExceLint.FeatureConf();

            // reference counts
            if (this.inDegree.Checked) { c = c.enableInDegree(); }
            if (this.outDegree.Checked) { c = c.enableOutDegree(); }
            if (this.combinedDegree.Checked) { c = c.enableCombinedDegree(); }

            // spatiostructual vectors
            if (this.inVectors.Checked) { c = c.enableShallowInputVectorMixedL2NormSum(); }
            if (this.outVectors.Checked) { c = c.enableShallowOutputVectorMixedL2NormSum(); }
            if (this.inVectorsAbs.Checked) { c = c.enableShallowInputVectorAbsoluteL2NormSum(); }
            if (this.outVectorsAbs.Checked) { c = c.enableShallowOutputVectorAbsoluteL2NormSum(); }

            // locality
            if (this.ProximityAbove.Checked) { c = c.enableProximityAbove(); }
            if (this.ProximityBelow.Checked) { c = c.enableProximityBelow(); }
            if (this.ProximityLeft.Checked) { c = c.enableProximityLeft(); }
            if (this.ProximityRight.Checked) { c = c.enableProximityRight(); }

            // Scopes (i.e., conditioned analysis)
            if (this.allCellsFreq.Checked) { c = c.analyzeRelativeToAllCells(); }
            if (this.columnCellsFreq.Checked) { c = c.analyzeRelativeToColumns(); }
            if (this.rowCellsFreq.Checked) { c = c.analyzeRelativeToRows(); }
            if (this.levelsFreq.Checked) { c = c.analyzeRelativeToLevels(); }

            // weighting / program resynthesis
            if (this.inferAddrModes.Checked) { c = c.inferAddressModes();  }
            if (!this.allCells.Checked) { c = c.analyzeOnlyFormulas();  }
            if (this.weightByIntrinsicAnomalousness.Checked) { c = c.weightByIntrinsicAnomalousness(); }

            return c;
        }

        #endregion UTILITY_FUNCTIONS

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
    }
}
