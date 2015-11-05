//The MIT License (MIT)

//Copyright (c) 20015 Jason Wilbur Turpin

//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the Software), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions

//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.

//THE SOFTWARE IS PROVIDED AS IS, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.

using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelExposer
{
    public partial class formMain : Form
    {
        private const string BUTTON_TEXT_DEFAULT = "Expose!";
        private const string BUTTON_TEXT_EXPOSING = "Exposing...";

        public formMain()
        {
            InitializeComponent();
        }

        private void buttonExpose_Click(object sender, EventArgs e)
        {
            buttonExpose.Text = BUTTON_TEXT_EXPOSING;
            this.Enabled = false;

            Excel.Application excelApp = null;

            Excel.Workbook sourceWorkbook = null;            

            excelApp = new Excel.Application();             
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;

            textBoxStatus.Text = string.Empty;

            UpdateStatus("Opening the Source Spreadsheet.");            
            
            sourceWorkbook = excelApp.Workbooks.Open(textBoxSourceFile.Text);

            if (sourceWorkbook.ProtectStructure)
            {
                UpdateStatus("Removing Strucural Protection");

                sourceWorkbook = UnprotectWorkbookStructure(sourceWorkbook);
            }

            int endRow, endColumn;

            int sourceSheetCount = sourceWorkbook.Sheets.Count;

            // overall paste, formulas are set afterwards
            for (int sourceSheetIndex = 1; sourceSheetIndex <= sourceWorkbook.Sheets.Count; sourceSheetIndex++)
            {
                Excel.Worksheet sourceSheet = sourceWorkbook.Sheets[sourceSheetIndex];

                UpdateStatus(string.Format("Removing Protection and Showing Sheet {0}/{1}, '{2}'", sourceSheetIndex, sourceSheetCount, sourceSheet.Name));

                if (sourceSheet.ProtectContents)
                {
                    sourceSheet = UnprotectSheet(sourceSheet);
                }

                if (sourceSheet.Visible == Excel.XlSheetVisibility.xlSheetHidden || sourceSheet.Visible == Excel.XlSheetVisibility.xlSheetVeryHidden)
                {
                    sourceSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                    sourceSheet.Tab.Color = Color.Red;
                }

                if (sourceSheet == null)
                {
                    MessageBox.Show(string.Format("Couldn't determine the password for the sheet '{0}'", sourceSheet.Name));
                }

                Excel.Range sourceUsedRange = sourceSheet.UsedRange;

                endRow = sourceUsedRange.Row + sourceUsedRange.Rows.Count;
                endColumn = sourceUsedRange.Column + sourceUsedRange.Columns.Count;

                Marshal.FinalReleaseComObject(sourceUsedRange);

                Excel.Range sourceCells = sourceSheet.Range[sourceSheet.Cells[1, 1], sourceSheet.Cells[endRow, endColumn]];

                // show all rows and columns?
                if (checkBoxShowHiddenRows.Checked)
                {
                    sourceCells.EntireRow.Hidden = false;
                }

                if (checkBoxShowHiddenColumns.Checked)
                {
                    sourceCells.EntireColumn.Hidden = false;
                }

                Marshal.FinalReleaseComObject(sourceCells);
                Marshal.FinalReleaseComObject(sourceSheet);
            }

            excelApp.Visible = true;
            excelApp.DisplayAlerts = true;
            Marshal.FinalReleaseComObject(excelApp);

            this.Enabled = true;
            buttonExpose.Text = BUTTON_TEXT_DEFAULT;
            UpdateStatus("Finished!");            

            MessageBox.Show("Finished, the Spreadsheet is Exposed!", "Excel Exposer!", MessageBoxButtons.OK);
        }

        private void textBoxSourceFile_DoubleClick(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                textBoxSourceFile.Text = openFileDialog.FileName;
            }

        }

        private static string _worksheetPassword = string.Empty;

        private Excel.Worksheet UnprotectSheet(Excel.Worksheet sheet)
        {
            if (!string.IsNullOrEmpty(_worksheetPassword))
            {
                try
                {
                    sheet.Unprotect(_worksheetPassword);
                }
                catch
                {

                }

                if (!sheet.ProtectContents)
                {                    
                    return sheet;
                }
            }

            // This code is based on a routine found here (public university, University of Wisconsin, Green Bay): https://uknowit.uwgb.edu/page.php?id=28850

            int attemptCount = 0;

            for (int i = 65; i < 67; i++)
                for (int j = 65; j < 67; j++)
                    for (int k = 65; k < 67; k++)
                        for (int l = 65; l < 67; l++)
                            for (int m = 65; m < 67; m++)
                                for (int i1 = 65; i1 < 67; i1++)
                                    for (int i2 = 65; i2 < 67; i2++)
                                        for (int i3 = 65; i3 < 67; i3++)
                                            for (int i4 = 65; i4 < 67; i4++)
                                                for (int i5 = 65; i5 < 67; i5++)
                                                    for (int i6 = 65; i6 < 67; i6++)
                                                        for (int n = 32; n < 126; n++)
                                                        {
                                                            _worksheetPassword = Convert.ToChar(i).ToString() + Convert.ToChar(j).ToString() + Convert.ToChar(k).ToString() + Convert.ToChar(l).ToString() + Convert.ToChar(m).ToString() +
                                                                Convert.ToChar(i1).ToString() + Convert.ToChar(i2).ToString() + Convert.ToChar(i3).ToString() + Convert.ToChar(i4).ToString() + Convert.ToChar(i5).ToString() +
                                                                Convert.ToChar(i6).ToString() + Convert.ToChar(n).ToString();

                                                            attemptCount++;

                                                            try
                                                            {
                                                                sheet.Unprotect(_worksheetPassword);
                                                            }
                                                            catch { }                                                          

                                                            if (!sheet.ProtectContents)
                                                            {
                                                                return sheet;
                                                            }
                                                        }

            return null;
        }

        private Excel.Workbook UnprotectWorkbookStructure(Excel.Workbook workbook)
        {
            string workbookStructurePassword = string.Empty;

            int attemptCount = 0;

            // This code is based on a routine found here (public university, University of Wisconsin, Green Bay): https://uknowit.uwgb.edu/page.php?id=28850

            for (int i = 65; i < 67; i++)
                for (int j = 65; j < 67; j++)
                    for (int k = 65; k < 67; k++)
                        for (int l = 65; l < 67; l++)
                            for (int m = 65; m < 67; m++)
                                for (int i1 = 65; i1 < 67; i1++)
                                    for (int i2 = 65; i2 < 67; i2++)
                                        for (int i3 = 65; i3 < 67; i3++)
                                            for (int i4 = 65; i4 < 67; i4++)
                                                for (int i5 = 65; i5 < 67; i5++)
                                                    for (int i6 = 65; i6 < 67; i6++)
                                                        for (int n = 32; n < 126; n++)
                                                        {
                                                            workbookStructurePassword = Convert.ToChar(i).ToString() + Convert.ToChar(j).ToString() + Convert.ToChar(k).ToString() + Convert.ToChar(l).ToString() + Convert.ToChar(m).ToString() +
                                                                Convert.ToChar(i1).ToString() + Convert.ToChar(i2).ToString() + Convert.ToChar(i3).ToString() + Convert.ToChar(i4).ToString() + Convert.ToChar(i5).ToString() +
                                                                Convert.ToChar(i6).ToString() + Convert.ToChar(n).ToString();

                                                            attemptCount++;

                                                            try
                                                            {
                                                                workbook.Unprotect(workbookStructurePassword);
                                                            }
                                                            catch { }
                                                            

                                                            if (!workbook.ProtectStructure)
                                                            {
                                                                return workbook;
                                                            }
                                                        }

            return null;
        }

        private void UpdateStatus(string status)
        {
            if (textBoxStatus.Text != string.Empty)
            {
                textBoxStatus.AppendText(Environment.NewLine);
            }
            
            textBoxStatus.AppendText(status);
            this.Refresh();
        }

    }
}
