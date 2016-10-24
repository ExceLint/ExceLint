using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.Text;
using System.Threading;

namespace ExceLintCLIGenerator
{
    public partial class Output : Form
    {
        bool cancelled = false;

        public Output(string exe, string[] flags)
        {
            InitializeComponent();

            Action<string> writeToWindow = (s) => outputWindow.AppendText(s);

            runCommand(exe, flags, writeToWindow);

            cancelled = true;

            cancelButton.Text = "Close";
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            if (!cancelled)
            {
                cancelled = true;

                cancelButton.Text = "Close";

            } else
            {
                Close();
            }
        }

        public void runCommand(string cpath, string[] args, Action<string> windowWriter)
        {
            using (var p = new Process())
            {
                // notice that we're using the Windows shell here and the unix-y 2>&1
                p.StartInfo.FileName = @"c:\windows\system32\cmd.exe";
                p.StartInfo.Arguments = "/c \"" + cpath + " " + String.Join(" ", args) + "\" 2>&1";
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.RedirectStandardError = true;

                using (var outputWaitHandle = new AutoResetEvent(false))
                {
                    p.OutputDataReceived += (sender, e) =>
                    {
                        // check for cancellation
                        if (cancelled)
                        {
                            if (e.Data != null)
                            {
                                windowWriter(e.Data);
                            }

                            windowWriter("Cancelled.");

                            return;
                        }

                        // attach event handler
                        if (e.Data == null)
                        {
                            outputWaitHandle.Set();
                        }
                        else
                        {
                            windowWriter(e.Data);
                        }
                    };

                    // start process
                    p.Start();

                    // begin async read
                    p.BeginOutputReadLine();

                    // wait for process to terminate
                    p.WaitForExit();

                    // wait on handle
                    outputWaitHandle.WaitOne();

                    // check exit code
                    if (p.ExitCode == 0)
                    {
                        windowWriter("Done.");
                    }
                }
            }
        }

        private void Output_Load(object sender, EventArgs e)
        {

        }
    }
}
