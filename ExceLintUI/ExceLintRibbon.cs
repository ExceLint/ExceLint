using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.FSharp.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Threading;
using ExceLintFileFormats;

namespace ExceLintUI
{
    public partial class ExceLintRibbon
    {
        private static bool USE_MULTITHREADED_UI = false;
        private static string DEFAULT_GROUND_TRUTH_FILENAME = "ground_truth";

        Dictionary<Excel.Workbook, WorkbookState> wbstates = new Dictionary<Excel.Workbook, WorkbookState>();
        System.Collections.Concurrent.ConcurrentDictionary<Excel.Workbook,Boolean> wbShutdown = new System.Collections.Concurrent.ConcurrentDictionary<Excel.Workbook,Boolean>();
        WorkbookState currentWorkbook;
        private ExceLintGroundTruth annotations;

        #region BUTTON_HANDLERS
        private void showVectors_Click(object sender, RibbonControlEventArgs e)
        {
            currentWorkbook.getMixedFormulaVectors(forceDAGBuild: forceBuildDAG.Checked);
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
                    wbs.flag(showFixes);
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
                } catch (Exception e) { }

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
            var time_str = "DAG construction ms: " + analysis.dag.AnalysisMilliseconds + "\n" +
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
                    // if, BEFORE the analysis, the user requests debug info
                    // AND the heatmap is PRESENTLY HIDDEN, show the debug info
                    // whether there ends up being a heatmap to show or not.
                    var debug_display = wbs.DebugMode && wbs.HeatMap_Hidden;

                    wbs.toggleHeatMap(WorkbookState.MAX_DURATION_IN_MS, conf, forceBuildDAG, pb);
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
                this.showFixes.Enabled = true;
            } else
            {
                this.allCellsFreq.Checked = true;
                this.rowCellsFreq.Checked = true;
                this.columnCellsFreq.Checked = true;
                this.levelsFreq.Checked = false;
                this.showFixes.Enabled = false;
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
                    rng.AddComment(annot.Item2.Comment);
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
            // get cursor location
            var cursor = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            AST.Address cursorAddr = ParcelCOMShim.Address.AddressFromCOMObject(cursor, Globals.ThisAddIn.Application.ActiveWorkbook);

            // get bug annotation from database
            var annot = Annotations.AnnotationFor(cursorAddr);

            // populate form and ask user for new data
            var mabf = new MarkAsBugForm(annot);

            // show form
            var result = mabf.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                // update database
                annot.BugKind = mabf.BugKind;
                annot.Note = mabf.Notes;
                Annotations.SetAnnotationFor(cursorAddr, annot);

                // stick note into workbook
                if (cursor.Comment != null)
                {
                    cursor.Comment.Delete();
                }
                cursor.AddComment(annot.Comment);
            }
        }

        private void ExportSquareMatrixButton_Click(object sender, RibbonControlEventArgs e)
        {
            var csv = CurrentWorkbook.GetSquareMatrices(forceBuildDAG.Checked, normRefCheckBox.Checked, normSSCheckBox.Checked);
            System.Windows.Forms.Clipboard.SetText(csv);
            System.Windows.Forms.MessageBox.Show("Exported to clipboard.");
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
            if (wbs == null || Globals.ThisAddIn.Application.ActiveProtectedViewWindow != null)
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
                this.weightByIntrinsicAnomalousness.Enabled = disabled;
                this.significanceTextBox.Enabled = disabled;
                this.conditioningSetSize.Enabled = disabled;
                this.spectralRanking.Enabled = disabled;
                this.showFixes.Enabled = disabled;
                this.annotate.Enabled = disabled;

                // tell the user ExceLint doesn't work
                SetTooltips(disabled_text);
            } else
            {
                // clear button text
                SetTooltips("");

                // enable auditing buttons if an audit has started
                this.MarkAsOKButton.Enabled = wbs.MarkAsOK_Enabled;
                this.StartOverButton.Enabled = wbs.ClearColoringButton_Enabled;
                this.AnalyzeButton.Enabled = wbs.Analyze_Enabled && wbs.HeatMap_Hidden;

                // only enable viewing heatmap if we are not in the middle of an analysis
                this.showHeatmap.Enabled = wbs.Analyze_Enabled;

                // disable config buttons if we are:
                // 1. in the middle of an audit, or
                // 2. we are viewing the heatmap, or
                // 3. if spectral ranking is checked, disable scopes
                var enable_config = wbs.Analyze_Enabled && wbs.HeatMap_Hidden;
                this.allCellsFreq.Enabled = enable_config && !this.spectralRanking.Checked;
                this.columnCellsFreq.Enabled = enable_config && !this.spectralRanking.Checked;
                this.rowCellsFreq.Enabled = enable_config && !this.spectralRanking.Checked;
                this.levelsFreq.Enabled = enable_config && !this.spectralRanking.Checked;
                this.DebugOutput.Enabled = enable_config;
                this.forceBuildDAG.Enabled = enable_config;
                this.inferAddrModes.Enabled = enable_config;
                this.allCells.Enabled = enable_config;
                this.weightByIntrinsicAnomalousness.Enabled = enable_config;
                this.significanceTextBox.Enabled = enable_config;
                this.conditioningSetSize.Enabled = enable_config;
                this.spectralRanking.Enabled = enable_config;
                this.showFixes.Enabled = enable_config && this.spectralRanking.Checked;

                // toggle the heatmap label depending on the heatmap shown/hidden state
                if (wbs.HeatMap_Hidden)
                {
                    this.showHeatmap.Label = "Show Heat Map";
                }
                else
                {
                    this.showHeatmap.Label = "Hide Heat Map";
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

            // reference counts
            if (this.inDegree.Checked) { c = c.enableInDegree(true); }
            if (this.outDegree.Checked) { c = c.enableOutDegree(true); }
            if (this.combinedDegree.Checked) { c = c.enableCombinedDegree(true); }

            // spatiostructual vectors
            if (this.inVectors.Checked) { c = c.enableShallowInputVectorMixedL2NormSum(true); }
            if (this.outVectors.Checked) { c = c.enableShallowOutputVectorMixedL2NormSum(true); }
            if (this.inVectorsAbs.Checked) { c = c.enableShallowInputVectorAbsoluteL2NormSum(true); }
            if (this.outVectorsAbs.Checked) { c = c.enableShallowOutputVectorAbsoluteL2NormSum(true); }

            // locality
            if (this.ProximityAbove.Checked) { c = c.enableProximityAbove(true); }
            if (this.ProximityBelow.Checked) { c = c.enableProximityBelow(true); }
            if (this.ProximityLeft.Checked) { c = c.enableProximityLeft(true); }
            if (this.ProximityRight.Checked) { c = c.enableProximityRight(true); }

            // Scopes (i.e., conditioned analysis)
            if (this.allCellsFreq.Checked) { c = c.analyzeRelativeToAllCells(true); }
            if (this.columnCellsFreq.Checked) { c = c.analyzeRelativeToColumns(true); }
            if (this.rowCellsFreq.Checked) { c = c.analyzeRelativeToRows(true); }
            if (this.levelsFreq.Checked) { c = c.analyzeRelativeToLevels(true); }

            // weighting / program resynthesis
            if (this.inferAddrModes.Checked) { c = c.inferAddressModes(true);  }
            if (!this.allCells.Checked) { c = c.analyzeOnlyFormulas(true);  }
            if (this.weightByIntrinsicAnomalousness.Checked) { c = c.weightByIntrinsicAnomalousness(true); }
            if (this.conditioningSetSize.Checked) { c = c.weightByConditioningSetSize(true); }

            // ranking type
            if (this.spectralRanking.Checked) { c = c.spectralRanking(true).analyzeRelativeToSheet(true); }

            // COF?
            if (this.useCOF.Checked) {
                c = c.enableShallowInputVectorMixedCOFRefUnnormSSNorm(true);
            } else
            {
                c = c.enableShallowInputVectorMixedCOFRefUnnormSSNorm(false);
            }

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
