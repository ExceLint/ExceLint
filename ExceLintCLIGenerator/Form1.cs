using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;
using System.Threading;

namespace ExceLintCLIGenerator
{


    public partial class Form1 : Form
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        private static extern int GetShortPathName(
            [MarshalAs(UnmanagedType.LPTStr)] string path,
            [MarshalAs(UnmanagedType.LPTStr)] StringBuilder shortPath,
            int shortPathLength
         );

        private static string shortPath(string p)
        {
            var sb = new StringBuilder(255);
            GetShortPathName(p, sb, sb.Capacity);
            return sb.ToString();
        }

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
            Properties.Settings.Default.Save();
        }

        private void excelintrunnerPathTextBox_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["excelintrunnerPathDefault"] = excelintrunnerPathTextBox.Text;
            Properties.Settings.Default.Save();
        }

        private void benchmarkDirTextbox_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["inputPathDefault"] = benchmarkDirTextbox.Text;
            Properties.Settings.Default.Save();
        }

        private void outputDirectoryTextbox_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["outputPathDefault"] = outputDirectoryTextbox.Text;
            Properties.Settings.Default.Save();
        }

        private void excelintGroundTruthCSVTextbox_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["excelintGTPathDefault"] = excelintGroundTruthCSVTextbox.Text;
            Properties.Settings.Default.Save();
        }

        private void custodesGroundTruthCSVTextbox_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["custodesGTPathDefault"] = custodesGroundTruthCSVTextbox.Text;
            Properties.Settings.Default.Save();
        }

        private void javaPathTextbox_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["javaPathDefault"] = javaPathTextbox.Text;
            Properties.Settings.Default.Save();
        }

        private void thresholdTextBox_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["thresholdDefault"] = Double.Parse(thresholdTextBox.Text);
            Properties.Settings.Default.Save();
        }

        private void custodesJARPathTextbox_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["custodesJARPathDefault"] = custodesJARPathTextbox.Text;
            Properties.Settings.Default.Save();
        }

        private void allcellsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["allCellsFlagDefault"] = allcellsCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        private void columnsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["columnsFlagDefault"] = columnsCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        private void rowsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["rowsFlagDefault"] = rowsCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        private void levelsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["levelsFlagDefault"] = levelsCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        private void sheetsCheckbox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["sheetsFlagDefault"] = sheetsCheckbox.Checked;
            Properties.Settings.Default.Save();
        }

        private void addrmodeCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["addrmodeFlagDefalt"] = addrmodeCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        private void weighIntrinsicCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["intrinsicFlagDefault"] = weighIntrinsicCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        private void weighCSSCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["cssFlagDefault"] = weighCSSCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        private void analyzeInputsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["inputsTooFlagDefault"] = analyzeInputsCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        private void verboseCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["verboseFlagDefault"] = verboseCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        private void noexitCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default["noexitFlagDefault"] = noexitCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        private string[] generateCLIInvocation(bool includeEXE)
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

            var flags = new List<string>();

            if (includeEXE) flags.Add(shortPath(excelintrunnerPathTextBox.Text));
            flags.Add(shortPath(benchmarkDirTextbox.Text));
            flags.Add(shortPath(outputDirectoryTextbox.Text));
            flags.Add(shortPath(excelintGroundTruthCSVTextbox.Text));
            flags.Add(shortPath(custodesGroundTruthCSVTextbox.Text));
            flags.Add(shortPath(javaPathTextbox.Text));
            flags.Add(shortPath(custodesJARPathTextbox.Text));
            flags.Add("-thresh " + thresholdTextBox.Text);

            if (verboseCheckBox.Checked) flags.Add("-verbose");
            if (noexitCheckBox.Checked) flags.Add("-noexit");
            if (spectralCheckBox.Checked) flags.Add("-spectral");
            if (allcellsCheckBox.Checked) flags.Add("-allcells");
            if (columnsCheckBox.Checked) flags.Add("-columns");
            if (rowsCheckBox.Checked) flags.Add("-rows");
            if (levelsCheckBox.Checked) flags.Add("-levels");
            if (sheetsCheckbox.Checked) flags.Add("-sheets");
            if (addrmodeCheckBox.Checked) flags.Add("-addrmode");
            if (weighIntrinsicCheckBox.Checked) flags.Add("-intrinsic");
            if (weighCSSCheckBox.Checked) flags.Add("-css");
            if (analyzeInputsCheckBox.Checked) flags.Add("-inputstoo");

            return flags.ToArray();
        }

        private string generateCLIInvocationString()
        {
            return String.Join(" ", generateCLIInvocation(includeEXE: true));
        }

        private void clipboardButton_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(generateCLIInvocationString());
        }

        public static void runCommand(string cpath, string[] args)
        {
            using (var p = new Process())
            {
                // notice that we're using the Windows shell here and the unix-y 2>&1
                p.StartInfo.FileName = @"c:\windows\system32\cmd.exe";
                p.StartInfo.Arguments = "/c \"" + cpath + " " + String.Join(" ", args) + "\" 2>&1";
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardOutput = false;
                p.StartInfo.RedirectStandardError = false;

                using (var outputWaitHandle = new AutoResetEvent(false))
                {
                    // start process
                    p.Start();

                    // wait for process to terminate
                    p.WaitForExit();

                    // wait on handle
                    outputWaitHandle.WaitOne();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var flags = generateCLIInvocation(includeEXE: false);
            runCommand(shortPath(excelintrunnerPathTextBox.Text), flags);
        }
    }
}
