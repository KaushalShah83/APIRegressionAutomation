using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace APIRegressionAutomation
{
    public partial class FrmAPITesting : Form
    {
        string file = "";
        public FrmAPITesting()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                file = openFileDialog1.FileName;
                txtInputFile.Text = file;
                
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(file))
            {
                this.Cursor = Cursors.WaitCursor;
                AutomateAPITesting class1 = new AutomateAPITesting();
                class1.ReadExcelFile(file, ref txtAPIResultOutput);
                this.Cursor = Cursors.Default;

            }
            else
                MessageBox.Show("Please select input file", "API Automation Testing");
        }

        private void txtAPIResultOutput_TextChanged(object sender, EventArgs e)
        {
            txtAPIResultOutput.SelectionStart = txtAPIResultOutput.Text.Length;
            txtAPIResultOutput.ScrollToCaret();
        }
    }
}
