// Copyright (c) 2012, Mark Zachacz and Michael Waite
//  All rights reserved.
//
// Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
// Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
// Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
//
// THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.


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
            if (outFile == null)
            {
                MessageBox.Show("Please select an output file.");
                return;
            }

            // file building utilities
            String delimiter = ",";
            StreamWriter oFile = new StreamWriter(outFile);
            String[] temp = new String[3];

            // write the first line in the .csv file
            temp[0] = "Cell";
            temp[1] = "Old Value";
            temp[2] = "New Value";
            oFile.WriteLine(String.Join(delimiter, temp));

            // old spreadsheet
            Object Item1 = mBook1SheetBox.SelectedItem;
            if (Item1 == null)
            {
                MessageBox.Show("Please select a worksheet for Workbook 1");
                return;
            }
            Excel._Worksheet sheet1 = (Excel._Worksheet)mBook1Sheets.Item[Item1.ToString()];
            
            // new spreadsheet
            Object Item2 = mBook2SheetBox.SelectedItem;
            if (Item1 == null)
            {
                MessageBox.Show("Please select a worksheet for Workbook 2");
                return;
            }
            Excel._Worksheet sheet2 = (Excel._Worksheet)mBook2Sheets.Item[Item2.ToString()];

            // Make the window wait.
            mDiff.Enabled = false;
            Cursor.Current = Cursors.WaitCursor;
            
            //create range objects for gathering the bounds of the spreadsheets
            Excel.Range range1 = sheet1.UsedRange;
            Excel.Range range2 = sheet2.UsedRange;

            // get the final filled row and column coordinate for sheet1
            int endRow1 = range1.Rows.Count;
            int endCol1 = range1.Columns.Count;

            // get the final filled row and column coordinate for sheet2
            int endRow2 = range2.Rows.Count;
            int endCol2 = range2.Columns.Count;

            // use the larger number for both rows and columns
            int endRow = (endRow1 > endRow2) ? endRow1 : endRow2;
            int endCol = (endCol1 > endCol2) ? endCol1 : endCol2;

            // diff the two files, and add a line onto the StringBuilder
            // if the cells are different
            for (int i = 1; i <= endCol; i++)
            {
                for (int j = 1; j <= endRow; j++)
                {
                    if (cellStr(sheet1.Cells[j, i]) != cellStr(sheet2.Cells[j, i]))
                    {
                        temp[0] = convert(j, i);
                        temp[1] = cellStr(sheet1.Cells[j, i]);
                        temp[2] = cellStr(sheet2.Cells[j, i]);
                        oFile.WriteLine(String.Join(delimiter, temp));
                    }
                }
            }

            // Close the file and make the window stop waiting.
            oFile.Close();
            mDiff.Enabled = true;
            Cursor.Current = Cursors.Default;
            
        }

        string cellStr(Excel.Range rhs)
        {
            if (rhs == null || rhs.Formula == null)
            {
                return "";
            }
            else
            {
                return "'" + rhs.Formula.ToString();
            }
        }


        // Converts a set of coordinates into an Excel coordinate.
        // Utility function for formatting output.
        // Obtained from StackOverflow.net, User: astander
        // http://stackoverflow.com/questions/1951517/convert-a-to-1-b-to-2-z-to-26-and-then-aa-to-27-ab-to-28-column-indexes-to
        // on October 27, 2012
        private String convert(int i, int j)
        {
            String columnString = "";
            decimal columnNumber = j;
            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }
            

            columnString += i.ToString();
            return columnString;
        }

        // Event handler for the top browse button.
        private void mBook1Load_Click(object sender, EventArgs e)
        {
            // Create a new OpenFileDialog object that will be used to select the first
            // workbook.
            System.Windows.Forms.OpenFileDialog fWin = new OpenFileDialog();

            // Set the extensions and default folder path to be used by the dialog.
            fWin.DefaultExt = "xlsx";
            fWin.Filter = "Spreadsheet Files (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv";
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

                if (fWin.FileName.Equals(mBook2File.Text) && fWin.FileName.EndsWith(".csv"))
                {
                    MessageBox.Show("Cannot compare a csv file to itself.");
                    return;
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
            fWin.Filter = "Spreadsheet Files (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv";
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

                if (fWin.FileName.Equals(mBook1File.Text) && fWin.FileName.EndsWith(".csv"))
                {
                    MessageBox.Show("Cannot compare a csv file to itself.");
                    return;
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

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Refrences to the first workbook to be diffed.
            mBook1Sheets = null;
            if (mBook1 != null)
            {
                mBook1.Close();
                mBook1 = null;
            }

            // Refrences to the second workbook to be diffed.
            mBook2Sheets = null;
            if (mBook2 != null)
            {
                mBook2.Close();
                mBook2 = null;
            }

            // Refrence to the instance of Microsoft Excel being used.
            mExcel.Quit();
            mExcel = null;
        }
    }
}
