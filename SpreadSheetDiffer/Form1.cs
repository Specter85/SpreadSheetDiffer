using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection; 

namespace SpreadSheetDiffer
{
    public partial class Form1 : Form
    {

        // Name of the CSV file to create.
        private String outFile = null;

        private Excel.Application mExcel = null;

        private Excel._Workbook mBook1 = null;
        private Excel.Sheets mBook1Sheets = null;

        private Excel._Workbook mBook2 = null;
        private Excel.Sheets mBook2Sheets = null;

        public Form1()
        {
            InitializeComponent();

            mExcel = new Excel.Application();
            mExcel.Visible = false;
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
                mBook1Sheets = null;
                if(mBook1 != null) {
                    mBook1.Close(false);
                    mBook1 = null;
                }

                mBook1File.Text = fWin.FileName;
                mBook1 = (Excel._Workbook)(mExcel.Workbooks.Add(fWin.FileName));
                mBook1Sheets = (Excel.Sheets)(mBook1.Sheets);

                mBook1Sheet.Enabled = true;
                mBook1Sheet.Items.Clear();

                int numSheets = mBook1Sheets.Count;
                for(int i = 1; i <= numSheets; i++) {
                    mBook1Sheet.Items.Add(((Excel._Worksheet)(mBook1Sheets.Item[i])).Name);
                }
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
                mBook2Sheets = null;
                if (mBook2 != null)
                {
                    mBook2.Close(false);
                    mBook2 = null;
                }

                mBook2File.Text = fWin.FileName;
                mBook2 = (Excel._Workbook)(mExcel.Workbooks.Add(fWin.FileName));
                mBook2Sheets = (Excel.Sheets)(mBook2.Sheets);

                mBook2Sheet.Enabled = true;
                mBook2Sheet.Items.Clear();

                int numSheets = mBook2Sheets.Count;
                for (int i = 1; i <= numSheets; i++)
                {
                    mBook2Sheet.Items.Add(((Excel._Worksheet)(mBook2Sheets.Item[i])).Name);
                }
            }
        }

        private void mCreate_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.SaveFileDialog fWin = new SaveFileDialog();

            fWin.DefaultExt = "xlsx";
            fWin.Filter = "CSV Files (*.csv)|*.csv";
            fWin.InitialDirectory =
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (fWin.ShowDialog() == DialogResult.OK)
            {
                outFile = fWin.FileName;
            }
        }
    }
}
