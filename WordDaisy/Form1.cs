using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordProcessingTest;

namespace WordDaisy {
    public partial class Form1 : Form {
        private string inputFile, outputDir;

        public Form1() {
            InitializeComponent();
        }

        private void openDocButton_Click(object sender, EventArgs e) {
            var res = openFileDialog1.ShowDialog();
            if (res == DialogResult.OK) {
                curDocNameLabel.Text = openFileDialog1.SafeFileName;
                inputFile = openFileDialog1.FileName;
            }
        }

        private void processDocButton_Click(object sender, EventArgs e) {
            var res = folderBrowserDialog1.ShowDialog();
            if (res == DialogResult.OK) {
                outputDir = folderBrowserDialog1.SelectedPath;
                folderBrowserDialog1.Dispose();

                Processor p = new Processor(inputFile, outputDir, openFileDialog1.SafeFileName);
                Boolean success = p.ProcessDocument();

                if (!success) {
                    doneLabel.Text = "Unable to open.";
                } else {
                    doneLabel.Text = "Done!";
                }
                doneLabel.Visible = true;
            }
        }
    }
}
