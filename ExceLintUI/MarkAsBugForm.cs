using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using ExceLintFileFormats;

namespace ExceLintUI
{
    public partial class MarkAsBugForm : Form
    {
        BugKind _selected_kind = BugKind.DefaultKind;
        string _notes = "";

        BugKind[] _sortedBugKinds = BugKind.AllKinds.OrderBy(bk => bk.ToString()).ToArray();
        Dictionary<BugKind, int> bkIndices = new Dictionary<BugKind, int>();

        public MarkAsBugForm(BugAnnotation[] bugAnnotations, AST.Address[] addrs)
        {
            StartPosition = FormStartPosition.CenterScreen;

            // get indices
            for (int i = 0; i < _sortedBugKinds.Length; i++)
            {
                bkIndices.Add(_sortedBugKinds[i], i);
            }
            // get indices
            int[] idxs = bugAnnotations.Select(ba => bkIndices[ba.BugKind]).ToArray();

            InitializeComponent();

            // populate combo box
            BugKindsCombo.DataSource = _sortedBugKinds.Select(bugkind => bugkind.ToString()).ToList();

            // is this a single annotation?
            bool singleAnnotation = bugAnnotations.Length == 1;

            if (singleAnnotation)
            {
                var bugAnnotation = bugAnnotations[0];
                var idx = idxs[0];
                var addr = addrs[0];

                // tell the user that we're doing a single annotation
                Text = "Annotate cell";
                markCellLabel.Text = "Mark cell " + addr.A1Local() + " as:";

                // select
                BugKindsCombo.SelectedIndex = idx;

                // fill notes
                bugNotesTextField.Text = bugAnnotation.Note;

                // fill checkboxes
                editBugKind.Checked = true;
                editNotes.Checked = true;

                // disable check controls
                editBugKind.Enabled = false;
                editNotes.Enabled = false;
            } else
            {
                // tell the user that we're doing multi-annotation
                Text = "Annotate multiple cells";
                markCellLabel.Text = "Mark multiple cells as:";

                // are all of the annotations the same?
                bool allSameBug = idxs.Distinct().Count() == 1;
                bool allSameNote = bugAnnotations.Select(ba => ba.Note).Distinct().Count() == 1;

                // if all bug kinds are the same, populate bug dropdown
                if (allSameBug)
                {
                    var annot_idx = idxs.Distinct().First();

                    // select combo box element
                    BugKindsCombo.SelectedIndex = annot_idx;

                    // fill check control
                    editBugKind.Checked = true;

                    // enable check control
                    editBugKind.Enabled = true;
                } else
                {
                    // fill check control
                    editBugKind.Checked = false;

                    // enable check control
                    editBugKind.Enabled = true;
                }

                // if all notes are the same, populate notes field
                if (allSameBug)
                {
                    var annot = bugAnnotations.Distinct().First();

                    // fill notes field
                    bugNotesTextField.Text = annot.Note;

                    // fill check control
                    editNotes.Checked = true;

                    // enable check control
                    editNotes.Enabled = true;
                }
                else
                {
                    // fill check control
                    editNotes.Checked = false;

                    // enable check control
                    editNotes.Enabled = true;
                }
            }
        }

        public BugKind BugKind
        {
            get
            {
                return _selected_kind;
            }
        }

        public string Notes
        {
            get
            {
                return _notes;
            }
        }

        private void markButton_Click(object sender, EventArgs e)
        {
            _selected_kind = _sortedBugKinds[BugKindsCombo.SelectedIndex];
            _notes = bugNotesTextField.Text;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
