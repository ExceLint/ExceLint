using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Depends;

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
        #endregion CONSTANTS

        private Excel.Application _app;
        private Excel.Workbook _workbook;
        private double _tool_significance = 0.95;
        private Dictionary<AST.Address, CellColor> _colors;
        private HashSet<AST.Address> _tool_highlights = new HashSet<AST.Address>();
        private HashSet<AST.Address> _output_highlights = new HashSet<AST.Address>();
        private HashSet<AST.Address> _known_good = new HashSet<AST.Address>();
        private IEnumerable<KeyValuePair<AST.Address, int>> _flaggable;
        private AST.Address _flagged_cell;
        private DAG _dag;
        private bool _debug_mode = false;

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

        public double toolProportion
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

        public void analyze(long max_duration_in_ms)
        {
            var sw = new System.Diagnostics.Stopwatch();
            sw.Start();

            using (var pb = new ProgBar())
            {
                // Disable screen updating during analysis to speed things up
                _app.ScreenUpdating = false;

                // Build dependency graph
                try
                {
                    _dag = new DAG(_app.ActiveWorkbook, _app, IGNORE_PARSE_ERRORS);
                }
                catch (Parcel.ParseException e)
                {
                    // cleanup UI and then rethrow
                    _app.ScreenUpdating = true;
                    throw e;
                }

                if (_dag.terminalInputVectors().Length == 0)
                {
                    System.Windows.Forms.MessageBox.Show("This spreadsheet contains no vector-input functions.");
                    _app.ScreenUpdating = true;
                    _flaggable = new KeyValuePair<AST.Address, int>[0];
                    return;
                }

                // TODO: DO WORK HERE
                KeyValuePair<AST.Address,int>[] scores;
                throw new NotImplementedException();

                if (_debug_mode)
                {
                    var score_str = String.Join("\n", scores.Take(10).Select(score => score.Key.A1FullyQualified() + " -> " + score.Value.ToString()));
                    System.Windows.Forms.MessageBox.Show(score_str);
                    System.Windows.Forms.Clipboard.SetText(score_str);
                }

                List<KeyValuePair<AST.Address, int>> high_scores = new List<KeyValuePair<AST.Address, int>>();

                // calculate cutoff idnex
                int thresh = scores.Length - Convert.ToInt32(scores.Length * _tool_significance);


                // filter out cells that are...
                _flaggable = scores.Where(pair => pair.Value >= scores[thresh].Value)   // below threshold
                                   .Where(pair => !_known_good.Contains(pair.Key))      // known to be good
                                   .Where(pair => pair.Value != 0).ToArray();           // score == 0

                // Enable screen updating when we're done
                _app.ScreenUpdating = true;

                sw.Stop();
            }
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
            _flaggable = _flaggable.Where(kvp => !_known_good.Contains(kvp.Key));
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
                if (_colors.ContainsKey(_flagged_cell))
                {
                    _colors[_flagged_cell] = cc;
                }
                else
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

        internal void fixError(Action<WorkbookState> setUIState)
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
                    analyze(MAX_DURATION_IN_MS);
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
    }
}
