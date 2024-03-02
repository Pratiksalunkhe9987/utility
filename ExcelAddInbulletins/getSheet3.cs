using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInbulletins
{
    internal class getSheet3
    {
        /* public void CopyDataToFinal_SD_Download_List()
         {
             Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

             if (workbook == null)
             {
                 MessageBox.Show("No active workbook found.");
                 return;
             }

             Excel.Worksheet sourceWorksheet = workbook.Sheets["Check Underscore & SD download"];
             Excel.Worksheet underscoreSDWorksheet = null;

             // Check if the underscoreSDWorksheet exists, if not, add it
             try
             {
                 underscoreSDWorksheet = workbook.Sheets["Final SD Download List"];
             }
             catch
             {
                 // underscoreSDWorksheet doesn't exist, so add it
                 underscoreSDWorksheet = (Excel.Worksheet)workbook.Sheets.Add();
                 underscoreSDWorksheet.Name = "Final SD Download List";
             }

             if (sourceWorksheet == null)
             {
                 MessageBox.Show("Source worksheet not found.");
                 return;
             }

             // Define the columns you want to copy (e.g., columns A, B, C)
             int[] columnsToCopy = { 1, 2, 3, 4, 5, 6, 7, 8 }; // Column indices are 1-based

             Excel.Range sourceRange = sourceWorksheet.UsedRange;
             Excel.Range destinationRange = underscoreSDWorksheet.Cells[1, 1];

             int destinationRow = 1;

             foreach (Excel.Range sourceRow in sourceRange.Rows)
             {
                 // Check if the row is marked in red
                 if (IsRowMarkedRed(sourceRow))
                 {
                     // Skip copying this row
                     continue;
                 }

                 // Copy data from the specified columns in the source row to the destination row
                 foreach (int column in columnsToCopy)
                 {
                     Excel.Range sourceCell = (Excel.Range)sourceRow.Cells[1, column];
                     Excel.Range destinationCell = (Excel.Range)destinationRange.Cells[destinationRow, column];

                     destinationCell.Value = sourceCell.Value;

                     sourceCell.Copy();
                     destinationCell.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
                     destinationCell.PasteSpecial(Excel.XlPasteType.xlPasteColumnWidths);
                     underscoreSDWorksheet.Application.CutCopyMode = Excel.XlCutCopyMode.xlCopy;
                 }

                 destinationRow++;
             }

             Excel.Range allCells = underscoreSDWorksheet.Cells;
             Excel.Borders borders = allCells.Borders;
             borders.LineStyle = Excel.XlLineStyle.xlContinuous;
             borders.Weight = Excel.XlBorderWeight.xlThin;

             MessageBox.Show("Data copied from Source to 'Check Underscore & SD download' sheet.");
         }
          // Helper method to check if a row is marked in red
          private bool IsRowMarkedRed(Excel.Range row)
          {
              return row.Interior.Color == System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
          }*/
        public void CopyHeaderToFinal_SD_Download_List()
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (workbook == null)
            {
                MessageBox.Show("No active workbook found.");
                return;
            }

            Excel.Worksheet sourceWorksheet = workbook.Sheets["Check Underscore & SD download"];
            Excel.Worksheet underscoreSDWorksheet = null;

            // Check if the underscoreSDWorksheet exists, if not, add it
            try
            {
                underscoreSDWorksheet = workbook.Sheets["Final SD Download List"];
            }
            catch
            {
                // underscoreSDWorksheet doesn't exist, so add it
                underscoreSDWorksheet = (Excel.Worksheet)workbook.Sheets.Add();
                underscoreSDWorksheet.Name = "Final SD Download List";
            }

            if (sourceWorksheet == null)
            {
                MessageBox.Show("Source worksheet not found.");
                return;
            }

            // Get the used range of the source worksheet
            Excel.Range sourceRange = sourceWorksheet.Rows[1]; // Only copy the first row (column headers)

            // Get the destination range (first row) in the destination worksheet
            Excel.Range destinationRange = underscoreSDWorksheet.Rows[1];

            // Copy the column headers from the source to the destination
            sourceRange.Copy(destinationRange);

            // Clear clipboard
            Clipboard.Clear();

            MessageBox.Show("Column headers copied from 'Check Underscore & SD download' sheet to 'Final SD Download List' sheet.");
        }
    }
}
