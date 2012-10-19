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

        // Refrence to the instance of Microsoft Excel being used.
        private Excel.Application mExcel = null;

        // Refrences to the first workbook to be diffed.
        private Excel._Workbook mBook1 = null;
        private Excel.Sheets mBook1Sheets = null;

        // Refrences to the second workbook to be diffed.
        private Excel._Workbook mBook2 = null;
        private Excel.Sheets mBook2Sheets = null;

        // Constructor.
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

        // Event handler for the top browse button.
        private void mBook1Load_Click(object sender, EventArgs e)
        {
            // Create a new OpenFileDialog object that will be used to select the first
            // workbook.
            System.Windows.Forms.OpenFileDialog fWin = new OpenFileDialog();

            // Set the extensions and default folder path to be used by the dialog.
            fWin.DefaultExt = "xlsx";
            fWin.Filter = "Excel Spreadsheets (*.xlsx)|*.xlsx";
            fWin.InitialDirectory = 
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // Open the dialog and do addional processing if a valid result is returned.
            if (fWin.ShowDialog() == DialogResult.OK)
            {
                // If a workbook is already open close it.
                mBook1Sheets = null;
                if(mBook1 != null) {
                    mBook1.Close(false);
                    mBook1 = null;
                }

                // Open the workbook selected in the file dialog.
                mBook1File.Text = fWin.FileName;
                mBook1 = (Excel._Workbook)(mExcel.Workbooks.Add(fWin.FileName));
                mBook1Sheets = (Excel.Sheets)(mBook1.Sheets);

                // Enable and clear the combobox for the first workbook's worksheets.
                mBook1SheetBox.Enabled = true;
                mBook1SheetBox.Items.Clear();

                // Populate the first workbook's combobox.
                int numSheets = mBook1Sheets.Count;
                for(int i = 1; i <= numSheets; i++) {
                    mBook1SheetBox.Items.Add(((Excel._Worksheet)(mBook1Sheets.Item[i])).Name);
                }
            }
        }

        // Event Handler for the middle browse button.
        private void mBook2Load_Click(object sender, EventArgs e)
        {
            // Create a new OpenFileDialog object that will be used to select the second
            // workbook.
            System.Windows.Forms.OpenFileDialog fWin = new OpenFileDialog();

            // Set the extensions and default folder path to be used by the dialog.
            fWin.DefaultExt = "xlsx";
            fWin.Filter = "Excel Spreadsheets (*.xlsx)|*.xlsx";
            fWin.InitialDirectory =
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // Open the dialog and do addional processing if a valid result is returned.
            if (fWin.ShowDialog() == DialogResult.OK)
            {
                // If a workbook is currently open close it.
                mBook2Sheets = null;
                if (mBook2 != null)
                {
                    mBook2.Close(false);
                    mBook2 = null;
                }

                // Open the workbook selected in the dialog.
                mBook2File.Text = fWin.FileName;
                mBook2 = (Excel._Workbook)(mExcel.Workbooks.Add(fWin.FileName));
                mBook2Sheets = (Excel.Sheets)(mBook2.Sheets);

                // Enable and clear the combobox for the second workbook's worksheets.
                mBook2SheetBox.Enabled = true;
                mBook2SheetBox.Items.Clear();

                // Populate the combobox for the second workbook.
                int numSheets = mBook2Sheets.Count;
                for (int i = 1; i <= numSheets; i++)
                {
                    mBook2SheetBox.Items.Add(((Excel._Worksheet)(mBook2Sheets.Item[i])).Name);
                }
            }
        }

        // Event handler for the bottom browse button.
        private void mCreate_Click(object sender, EventArgs e)
        {
            // Create a new SaveFileDialog object.
            System.Windows.Forms.SaveFileDialog fWin = new SaveFileDialog();

            // Set the extensions and default folder path to be used by the dialog.
            fWin.DefaultExt = "csv";
            fWin.Filter = "CSV Files (*.csv)|*.csv";
            fWin.InitialDirectory =
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // Set variables based on the result of the dialog. If an invalid result is
            // returned do nothing.
            if (fWin.ShowDialog() == DialogResult.OK)
            {
                outFile = fWin.FileName;
                mOutFileName.Text = fWin.FileName;
            }
        }
    }
}
