using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;

using Excel = Microsoft.Office.Interop.Excel;

namespace Inspector
{
    public partial class Inspector : Form
    {
        private string input;
        private string output;
        private string mappingConfFile;

        public Inspector()
        {
            InitializeComponent();
            this.MaximizeBox = false;
            this.MinimizeBox = false;
        }

        private void tbInput_Click(object sender, EventArgs e)
        {
            tbInput.Text = ShowFileSelectorDialog();
        }

        private void tbOutput_Click(object sender, EventArgs e)
        {
            this.tbOutput.Text = ShowFileSelectorDialog();
        }

        private void tbMappingConf_Click(object sender, EventArgs e)
        {
            this.tbMappingConf.Text = ShowOneFileSelectorDialog();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            this.input = this.tbInput.Text;
            this.output = this.tbOutput.Text;
            this.mappingConfFile = this.tbMappingConf.Text;

            if (input == string.Empty || output == string.Empty)
            {
                MessageBox.Show("Input and output path needed!");
                return;
            }

            this.btnStart.Enabled = false;

            Thread thread = new Thread(new ThreadStart(doCheck));
            thread.Start();
        }

        public void doCheck()
        {
            string outputFile = this.output + "\\result.xlsx";
            var writer = new ExcelWriter(outputFile);
            writer.init();

            string[] files = GetAllFilesRecursively(this.input);

            for (int i = 0; i < files.Length; i++)
            {
                Checker checker = new Checker(files[i], writer, this.mappingConfFile);
                string log = checker.Process();

                this.Invoke((MethodInvoker)delegate
                {
                    logArea.AppendText(log); // runs on UI thread
                });
            }

            writer.SaveAndClose();

            this.Invoke((MethodInvoker)delegate
            {
                this.btnStart.Enabled = true; // runs on UI thread
                MessageBox.Show("Done!", "Inspector");
            });
        }

        private string ShowFileSelectorDialog()
        {
            string foldername = string.Empty;

            FolderBrowserDialog dialog = new FolderBrowserDialog();
            //dialog.ShowNewFolderButton = false;
            //dialog.RootFolder = System.Environment.SpecialFolder.MyComputer;
            DialogResult result = dialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                foldername = dialog.SelectedPath;
            }

            return foldername;
        }

        private string ShowOneFileSelectorDialog()
        {
            string filename = string.Empty;

            var dialog = new OpenFileDialog();
            dialog.Filter = "TXT files|*.txt";
            DialogResult result = dialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                filename = dialog.FileName;
            }

            return filename;
        }

        private string[] GetAllFilesRecursively(string path)
        {
            string[] files = Directory.GetFiles(path, "*.doc", SearchOption.AllDirectories);
            return files.Where(f =>
                {
                    var attr = File.GetAttributes(f);
                    return (attr & FileAttributes.Hidden) != FileAttributes.Hidden;
                }).ToArray();
        }
    }
}
