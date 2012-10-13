using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SpreadSheetDiffer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // This is were the diffing code should go/start.
        private void mDiff_Click(object sender, EventArgs e)
        {

        }

        private void mBook1Load_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog fWin = new OpenFileDialog();

            fWin.DefaultExt = "xlsx";
            fWin.Filter = "Excel Spreadsheets (*.xlsx)|*.xlsx";
            fWin.InitialDirectory = 
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (fWin.ShowDialog() == DialogResult.OK)
            {
            }
        }

        private void mBook2Load_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog fWin = new OpenFileDialog();

            fWin.DefaultExt = "xlsx";
            fWin.Filter = "Excel Spreadsheets (*.xlsx)|*.xlsx";
            fWin.InitialDirectory =
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (fWin.ShowDialog() == DialogResult.OK)
            {
            }
        }

        private void mCreate_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.SaveFileDialog fWin = new OpenFileDialog();

            fWin.DefaultExt = "xlsx";
            fWin.Filter = "Excel Spreadsheets (*.xlsx)|*.xlsx";
            fWin.InitialDirectory =
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (fWin.ShowDialog() == DialogResult.OK)
            {
            }
        }
    }
}
