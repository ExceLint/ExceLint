using System;
using System.Windows.Forms;

namespace ExceLintCLIGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // read defaults
            excelintrunnerPathTextBox.Text = (string)Properties.Settings.Default["excelintrunnerPathDefault"];
            benchmarkDirTextbox.Text = (string)Properties.Settings.Default["inputPathDefault"];
            outputDirectoryTextbox.Text = (string)Properties.Settings.Default["outputPathDefault"];
            excelintGroundTruthCSVTextbox.Text = (string)Properties.Settings.Default["excelintGTPathDefault"];
            custodesGroundTruthCSVTextbox.Text = (string)Properties.Settings.Default["custodesGTPathDefault"];
            custodesJARPathTextbox.Text = (string)Properties.Settings.Default["custodesJARPathDefault"];
            javaPathTextbox.Text = (string)Properties.Settings.Default["javaPathDefault"];
            thresholdTextBox.Text = Convert.ToString((double)Properties.Settings.Default["thresholdDefault"]);
            verboseCheckBox.Checked = (bool)Properties.Settings.Default["verboseFlagDefault"];
            noexitCheckBox.Checked = (bool)Properties.Settings.Default["noexitFlagDefault"];
            spectralCheckBox.Checked = (bool)Properties.Settings.Default["spectralFlagDefault"];
            allcellsCheckBox.Checked = (bool)Properties.Settings.Default["allCellsFlagDefault"];
            columnsCheckBox.Checked = (bool)Properties.Settings.Default["columnsFlagDefault"];
            rowsCheckBox.Checked = (bool)Properties.Settings.Default["rowsFlagDefault"];
            levelsCheckBox.Checked = (bool)Properties.Settings.Default["levelsFlagDefault"];
            sheetsCheckbox.Checked = (bool)Properties.Settings.Default["sheetsFlagDefault"];
            weighIntrinsicCheckBox.Checked = (bool)Properties.Settings.Default["intrinsicFlagDefault"];
            weighCSSCheckBox.Checked = (bool)Properties.Settings.Default["cssFlagDefault"];
            analyzeInputsCheckBox.Checked = (bool)Properties.Settings.Default["inputsTooFlagDefault"];
            addrmodeCheckBox.Checked = (bool)Properties.Settings.Default["addrmodeFlagDefalt"];
        }

        private void spectralCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (spectralCheckBox.Checked)
            {
                sheetsCheckbox.Checked = true;
                allcellsCheckBox.Checked = false;
                columnsCheckBox.Checked = false;
                rowsCheckBox.Checked = false;
                levelsCheckBox.Checked = false;

                sheetsCheckbox.Enabled = false;
                allcellsCheckBox.Enabled = false;
                columnsCheckBox.Enabled = false;
                rowsCheckBox.Enabled = false;
                levelsCheckBox.Enabled = false;
            } else
            {
                sheetsCheckbox.Enabled = true;
                allcellsCheckBox.Enabled = true;
                columnsCheckBox.Enabled = true;
                rowsCheckBox.Enabled = true;
                levelsCheckBox.Enabled = true;
            }

            Properties.Settings.Default["spectralFlagDefault"] = spectralCheckBox.Checked;
            Properties.Settings.Default["allCellsFlagDefault"] = allcellsCheckBox.Checked;
            Properties.Settings.Default["columnsFlagDefault"] = columnsCheckBox.Checked;
            Properties.Settings.Default["rowsFlagDefault"] = rowsCheckBox.Checked;
            Properties.Settings.Default["levelsFlagDefault"] = levelsCheckBox.Checked;
            Properties.Settings.Default["sheetsFlagDefault"] = sheetsCheckbox.Checked;
        }

        private void excelintrunnerPathTextBox_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["excelintrunnerPathDefault"] = excelintrunnerPathTextBox.Text;
        }

        private void benchmarkDirTextbox_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["inputPathDefault"] = benchmarkDirTextbox.Text;
        }

        private void outputDirectoryTextbox_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["outputPathDefault"] = outputDirectoryTextbox.Text;
        }

        private void excelintGroundTruthCSVTextbox_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["excelintGTPathDefault"] = excelintGroundTruthCSVTextbox.Text;
        }

        private void custodesGroundTruthCSVTextbox_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["custodesGTPathDefault"] = custodesGroundTruthCSVTextbox.Text;
        }

        private void javaPathTextbox_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["javaPathDefault"] = javaPathTextbox.Text;
        }

        private void thresholdTextBox_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["thresholdDefault"] = Double.Parse(thresholdTextBox.Text);
        }

        private void custodesJARPathTextbox_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["custodesJARPathDefault"] = custodesJARPathTextbox.Text;
        }

        private void allcellsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["allCellsFlagDefault"] = allcellsCheckBox.Checked;
        }

        private void columnsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["columnsFlagDefault"] = columnsCheckBox.Checked;
        }

        private void rowsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["rowsFlagDefault"] = rowsCheckBox.Checked;
        }

        private void levelsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["levelsFlagDefault"] = levelsCheckBox.Checked;
        }

        private void sheetsCheckbox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["sheetsFlagDefault"] = sheetsCheckbox.Checked;
        }

        private void addrmodeCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["addrmodeFlagDefalt"] = addrmodeCheckBox.Checked;
        }

        private void weighIntrinsicCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["intrinsicFlagDefault"] = weighIntrinsicCheckBox.Checked;
        }

        private void weighCSSCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["cssFlagDefault"] = weighCSSCheckBox.Checked;
        }

        private void analyzeInputsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["inputsTooFlagDefault"] = analyzeInputsCheckBox.Checked;
        }

        private void verboseCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["verboseFlagDefault"] = verboseCheckBox.Checked;
        }

        private void noexitCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["noexitFlagDefault"] = noexitCheckBox.Checked;
        }

        private string generateCLIInvocation()
        {
            /* 
            -verbose    log per-spreadsheet flagged cells as separate CSVs
            -noexit     prompt user to press a key before exiting
            -spectral   use spectral outliers, otherwise use summation outliers;
                        forces the use of -sheets below and disables -allcells,
                        -columns, -rows, and -levels
            -allcells   condition by all cells
            -columns    condition by columns
            -rows       condition by rows
            -levels     condition by levels
            -sheets     condition by sheets
            -addrmode   infer address modes
            -intrinsic  weigh by intrinsic anomalousness
            -css        weigh by conditioning set size
            -inputstoo  analyze inputs as well; by default ExceLint only
                        analyzes formulas
            -thresh <n> sets max % to inspect at n%; default 5%
            */

            var flags =
                verboseCheckBox.Checked ? "-verbose" : "" +
                noexitCheckBox.Checked ? "-"

            return
                excelintrunnerPathTextBox.Text + " " +
                benchmarkDirTextbox.Text + " " + 
                outputDirectoryTextbox.Text + " " +
                excelintGroundTruthCSVTextbox.Text + " " +
                custodesGroundTruthCSVTextbox.Text + " " +
                javaPathTextbox.Text + " " +
                custodesJARPathTextbox.Text + " " +
                thresholdTextBox.Text + " " +
                flags
                ;
        }

        private void clipboardButton_Click(object sender, EventArgs e)
        {

        }
    }
}
