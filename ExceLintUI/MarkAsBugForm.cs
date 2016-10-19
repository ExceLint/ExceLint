using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using ExceLintFileFormats;

namespace ExceLintUI
{
    public struct BugAnnotation
    {
        public BugKind bugkind;
        public string notes;
    };

    public partial class MarkAsBugForm : Form
    {
        AST.Address _cell;
        Dictionary<AST.Address, BugAnnotation> _ba;
        BugKind _selected_kind = BugKind.DefaultKind;
        string _notes = "";

        BugKind[] _sortedBugKinds = BugKind.AllKinds.OrderBy(bk => bk.ToString()).ToArray();
        Dictionary<BugKind, int> bkIndices = new Dictionary<BugKind, int>();

        public MarkAsBugForm(AST.Address cell, Dictionary<AST.Address,BugAnnotation> bugAnnotations)
        {
            _cell = cell;
            _ba = bugAnnotations;

            // populate combo box
            BugKindsCombo.DataSource = _sortedBugKinds;

            // get indices
            for (int i = 0; i < _sortedBugKinds.Length; i++)
            {
                bkIndices.Add(_sortedBugKinds[i], i);
            }
            
            // if the user has already annotated, then...
            if (bugAnnotations.ContainsKey(_cell))
            {
                var bk = bugAnnotations[_cell].bugkind;
                var notes = bugAnnotations[_cell].notes;

                // get index 
                var idx = bkIndices[bk];

                // select combo box element
                BugKindsCombo.SelectedIndex = idx;

                // fill notes
                bugNotesTextField.Text = notes;
            } else
            {
                BugKindsCombo.SelectedIndex = bkIndices[BugKind.NotABug];
            }

            InitializeComponent();
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
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {

        }
    }
}
