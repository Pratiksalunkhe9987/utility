using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace ExcelAddInbulletins
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            MessageBox.Show("Ribbon Loaded");
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (workbook == null)
            {
                MessageBox.Show("No active workbook found.");
                return;
            }

            Excel.Worksheet sourceWorksheet = workbook.Sheets["Initial"];
            Excel.Worksheet destinationWorksheet = null;

            // Check if Sheet2 exists, if not, add it
            try
            {
                destinationWorksheet = workbook.Sheets["Remove_Redundant_Column "];
            }
            catch
            {
                // Sheet2 doesn't exist, so add it
                destinationWorksheet = (Excel.Worksheet)workbook.Sheets.Add();
                destinationWorksheet.Name = "Remove_Redundant_Column ";
            }

            if (sourceWorksheet == null)
            {
                MessageBox.Show("Source worksheet not found.");
                return;
            }

            // Define the columns you want to copy (e.g., columns A, B, C)
            int[] columnsToCopy = { 1, 2, 3, 4, 5, 6 }; // Column indices are 1-based
            int[] additionalColumns = { columnsToCopy.Length + 1, columnsToCopy.Length + 2 }; // Indices for additional columns

            // Add headers for additional columns in destination worksheet
            destinationWorksheet.Cells[1, additionalColumns[0]].Value = "Enhanacment(Y/N)";
            destinationWorksheet.Cells[1, additionalColumns[1]].Value = "Remarks";

            Excel.Range sourceRange = sourceWorksheet.UsedRange;
            Excel.Range destinationRange = destinationWorksheet.Cells[1, 1];

            int destinationRow = 1;

            foreach (Excel.Range sourceRow in sourceRange.Rows)
            {
                // Copy data from the specified columns in the source row to the destination row
                foreach (int column in columnsToCopy)
                {
                    Excel.Range sourceCell = (Excel.Range)sourceRow.Cells[1, column];
                    Excel.Range destinationCell = (Excel.Range)destinationRange.Cells[destinationRow, column];

                    destinationCell.Value = sourceCell.Value;

                    sourceCell.Copy();
                    destinationCell.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
                    destinationCell.PasteSpecial(Excel.XlPasteType.xlPasteColumnWidths);
                    destinationWorksheet.Application.CutCopyMode = Excel.XlCutCopyMode.xlCopy;
                }

                destinationRow++;
            }

            Excel.Range allCells = destinationWorksheet.Cells;
            Excel.Borders borders = allCells.Borders;
            borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            borders.Weight = Excel.XlBorderWeight.xlThin;

            //sourceRange.Copy(destinationRange);
            MessageBox.Show("Data copied from Source to Destination.");
            // MessageBox.Show("Data copied from Sheet1 to Sheet2.");

           // getSheet2 obj = new getSheet2();
           // obj.CopyDataToUnderscoreSDWorksheet();
        }
     
        private void button2_Click_1(object sender, RibbonControlEventArgs e)
        {
            FileZillaClient obj = new FileZillaClient();
          //  obj.DownloadFile(new string[] { "100010", "100014", "100085" });
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelReader reader = new ExcelReader();
            List<string> issueIds = reader.ReadIssueIds();
            foreach (string id in issueIds)
            {
                Console.WriteLine(id);
            }

            FileZillaClient obj = new FileZillaClient();
            //obj.DownloadFile(issueIds);

            foreach (string id in issueIds)
            {
                obj.Download(id);
               //obj.Downloadvariant(id);
            }

            // Output a message indicating that the files have been downloaded
            Console.WriteLine("All files downloaded successfully.");
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
             getSheet2 obj = new getSheet2();
             obj.CopyDataToUnderscoreSDWorksheet();
        }
    }
}
