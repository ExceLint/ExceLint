using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExceLintUI
{
    public struct BugAnnotation
    {
        public ExceLint.BugKind bugkind;
        public string notes;
    };

    public partial class MarkAsBugForm : Form
    {
        AST.Address _cell;
        Dictionary<AST.Address, BugAnnotation> _ba;
        ExceLint.BugKind _selected_kind = ExceLint.BugKind.DefaultKind;
        string _notes = "";

        ExceLint.BugKind[] _sortedBugKinds = ExceLint.BugKind.AllKinds.OrderBy(bk => bk.ToString()).ToArray();
        Dictionary<ExceLint.BugKind, int> bkIndices = new Dictionary<ExceLint.BugKind, int>();

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
                BugKindsCombo.SelectedIndex = bkIndices[ExceLint.BugKind.NotABug];
            }

            InitializeComponent();
        }

        public ExceLint.BugKind BugKind
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
