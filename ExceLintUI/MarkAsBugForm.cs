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
        BugAnnotation _ba;
        BugKind _selected_kind = BugKind.DefaultKind;
        string _notes = "";

        BugKind[] _sortedBugKinds = BugKind.AllKinds.OrderBy(bk => bk.ToString()).ToArray();
        Dictionary<BugKind, int> bkIndices = new Dictionary<BugKind, int>();

        public MarkAsBugForm(BugAnnotation bugAnnotation)
        {
            StartPosition = FormStartPosition.CenterScreen;

            // get indices
            for (int i = 0; i < _sortedBugKinds.Length; i++)
            {
                bkIndices.Add(_sortedBugKinds[i], i);
            }

            // get index 
            var idx = bkIndices[bugAnnotation.BugKind];

            InitializeComponent();

            // populate combo box
            BugKindsCombo.DataSource = _sortedBugKinds.Select(bugkind => bugkind.ToString()).ToList();

            // select combo box element
            BugKindsCombo.SelectedIndex = idx;

            // fill notes
            bugNotesTextField.Text = bugAnnotation.Note;
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
