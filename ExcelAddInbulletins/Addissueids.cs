using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInbulletins
{
    internal class Addissueids
    {
        public void AddIds(List<string> issueIds)
        {
         // Sort the issue IDs
            issueIds.Sort();

            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (workbook == null)
            {
                MessageBox.Show("No active workbook found.");
                return;
            }

            Excel.Worksheet underscoreSDWorksheet = null;

            try
            {
                underscoreSDWorksheet = workbook.Sheets["Final SD Download List"];
            }
            catch
            {
                MessageBox.Show("Final SD Download List sheet not found.");
                return;
            }

            // Get the index of the "Issue ID" column
            int issueIdColumnIndex = GetIssueIdColumnIndex(underscoreSDWorksheet);

            if (issueIdColumnIndex == -1)
            {
                MessageBox.Show("Issue ID column not found.");
                return;
            }

            // Find the last used row in the "Issue ID" column
            int lastRow = underscoreSDWorksheet.Cells[underscoreSDWorksheet.Rows.Count, issueIdColumnIndex].End[Excel.XlDirection.xlUp].Row;

            // Start adding IDs from the next row
            int startRow = lastRow + 1;

            // Write issue IDs to the "Final SD Download List" sheet
            for (int i = 0; i < issueIds.Count; i++)
            {
                Excel.Range cell = (Excel.Range)underscoreSDWorksheet.Cells[startRow + i, issueIdColumnIndex];
                cell.Value = issueIds[i];
            }

            MessageBox.Show($"Sorted and added issue IDs to the 'Final SD Download List' sheet.");
        }

        // Helper method to get the index of the "Issue ID" column
        private int GetIssueIdColumnIndex(Excel.Worksheet worksheet)
        {
            int columnIndex = -1;
            Excel.Range firstRow = worksheet.Rows[1];

            foreach (Excel.Range cell in firstRow.Cells)
            {
                if (cell.Value != null && cell.Value.ToString() == "IssueID")
                {
                    columnIndex = cell.Column;
                    break;
                }
            }

            return columnIndex;
        }

        /*public void CopyDataToFinal_SD_Download_List()
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (workbook == null)
            {
                MessageBox.Show("No active workbook found.");
                return;
            }

            Excel.Worksheet sourceWorksheet = workbook.Sheets["Check Underscore & SD download"];
            Excel.Worksheet finalSDWorksheet = null;

            try
            {
                finalSDWorksheet = workbook.Sheets["Final SD Download List"];
            }
            catch
            {
                finalSDWorksheet = (Excel.Worksheet)workbook.Sheets.Add();
                finalSDWorksheet.Name = "Final SD Download List";
            }

            if (sourceWorksheet == null)
            {
                MessageBox.Show("Source worksheet not found.");
                return;
            }

            // Get the used range of the source worksheet
            Excel.Range sourceRange = sourceWorksheet.UsedRange;

            // Get the last used row in the final worksheet
            int lastFinalRow = finalSDWorksheet.Cells[finalSDWorksheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

            // Iterate over each row in the source worksheet
            foreach (Excel.Range sourceRow in sourceRange.Rows)
            {
                string sourceIssueId = sourceRow.Cells[1, 1].Value?.ToString(); // Assuming IssueID is in the first column

                // Iterate over each row in the final worksheet
                for (int i = 2; i <= lastFinalRow; i++) // Start from row 2 to skip the header row
                {
                    string finalIssueId = finalSDWorksheet.Cells[i, 1].Value?.ToString(); // Assuming IssueID is in the first column

                    // Check if the source issue ID is contained within the final issue ID or is a variant
                    if (!string.IsNullOrEmpty(finalIssueId) && finalIssueId.Contains(sourceIssueId))
                    {
                        // Copy all other column data from the source row to the final row
                        for (int j = 1; j <= sourceWorksheet.Columns.Count; j++)
                        {
                            finalSDWorksheet.Cells[i, j].Value = sourceRow.Cells[1, j].Value;
                        }
                        break; // Stop searching for a match once a match is found
                    }
                }
            }

            MessageBox.Show("Data copied to 'Final SD Download List' sheet based on matching issue IDs.");
        }*/
        /* public void CopyDataToFinal_SD_Download_List()
         {
             Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

             if (workbook == null)
             {
                 MessageBox.Show("No active workbook found.");
                 return;
             }

             Excel.Worksheet sourceWorksheet = workbook.Sheets["Check Underscore & SD download"];
             Excel.Worksheet finalSDWorksheet = null;

             try
             {
                 finalSDWorksheet = workbook.Sheets["Final SD Download List"];
             }
             catch
             {
                 finalSDWorksheet = (Excel.Worksheet)workbook.Sheets.Add();
                 finalSDWorksheet.Name = "Final SD Download List";
             }

             if (sourceWorksheet == null)
             {
                 MessageBox.Show("Source worksheet not found.");
                 return;
             }

             // Get the used range of the source worksheet
             Excel.Range sourceRange = sourceWorksheet.UsedRange;

             // Get the last used row in the final worksheet
             int lastFinalRow = finalSDWorksheet.Cells[finalSDWorksheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

             // Iterate over each row in the source worksheet
             foreach (Excel.Range sourceRow in sourceRange.Rows)
             {
                 string sourceIssueId = sourceRow.Cells[1, 1]?.Value?.ToString(); // Assuming IssueID is in the first column

                 if (string.IsNullOrEmpty(sourceIssueId))
                 {
                     continue; // Skip processing if IssueID is null or empty
                 }

                 // Iterate over each row in the final worksheet
                 for (int i = 2; i <= lastFinalRow; i++) // Start from row 2 to skip the header row
                 {
                     string finalIssueId = finalSDWorksheet.Cells[i, 1]?.Value?.ToString(); // Assuming IssueID is in the first column

                     // Check if the source issue ID is contained within the final issue ID or is a variant
                     if (!string.IsNullOrEmpty(finalIssueId) && finalIssueId.Contains(sourceIssueId))
                     {
                         // Copy all other column data from the source row to the final row
                         for (int j = 1; j <= sourceWorksheet.Columns.Count; j++)
                         {
                             finalSDWorksheet.Cells[i, j].Value = sourceRow.Cells[1, j].Value;
                         }
                         break; // Stop searching for a match once a match is found
                     }
                 }
             }

             MessageBox.Show("Data copied to 'Final SD Download List' sheet based on matching issue IDs.");
         }*/
        /* public void CopyDataToFinal_SD_Download_List()
         {
             Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

             if (workbook == null)
             {
                 MessageBox.Show("No active workbook found.");
                 return;
             }

             Excel.Worksheet sourceWorksheet = workbook.Sheets["Check Underscore & SD download"];
             Excel.Worksheet finalSDWorksheet = null;

             try
             {
                 finalSDWorksheet = workbook.Sheets["Final SD Download List"];
             }
             catch
             {
                 finalSDWorksheet = (Excel.Worksheet)workbook.Sheets.Add();
                 finalSDWorksheet.Name = "Final SD Download List";
             }

             if (sourceWorksheet == null)
             {
                 MessageBox.Show("Source worksheet not found.");
                 return;
             }

             // Get the used range of the source worksheet
             Excel.Range sourceRange = sourceWorksheet.UsedRange;

             // Get the last used row in the final worksheet
             int lastFinalRow = finalSDWorksheet.Cells[finalSDWorksheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

             // Iterate over each row in the source worksheet
             foreach (Excel.Range sourceRow in sourceRange.Rows)
             {
                 string sourceIssueId = sourceRow.Cells[1, 1]?.Value?.ToString(); // Assuming IssueID is in the first column

                 if (string.IsNullOrEmpty(sourceIssueId))
                 {
                     continue; // Skip processing if IssueID is null or empty
                 }

                 // Iterate over each row in the final worksheet
                 for (int i = 2; i <= lastFinalRow; i++) // Start from row 2 to skip the header row
                 {
                     string finalIssueId = finalSDWorksheet.Cells[i, 1]?.Value?.ToString(); // Assuming IssueID is in the first column

                     // Check if the source issue ID is contained within the final issue ID or is a variant
                     if (!string.IsNullOrEmpty(finalIssueId) && (finalIssueId == sourceIssueId || finalIssueId.StartsWith(sourceIssueId + "_")))
                     {
                         // Copy all other column data from the source row to the final row, excluding the IssueID column
                         for (int j = 2; j <= sourceWorksheet.Columns.Count; j++) // Start from column 2 to skip the IssueID column
                         {
                             finalSDWorksheet.Cells[i, j].Value = sourceRow.Cells[1, j].Value;
                         }
                         break; // Stop searching for a match once a match is found
                     }
                 }
             }

             MessageBox.Show("Data copied to 'Final SD Download List' sheet based on matching issue IDs.");
         }
         */
        public void CopyDataToFinal_SD_Download_List()
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (workbook == null)
            {
                MessageBox.Show("No active workbook found.");
                return;
            }

            Excel.Worksheet sourceWorksheet = workbook.Sheets["Check Underscore & SD download"];
            Excel.Worksheet finalSDWorksheet = null;

            try
            {
                finalSDWorksheet = workbook.Sheets["Final SD Download List"];
            }
            catch
            {
                finalSDWorksheet = (Excel.Worksheet)workbook.Sheets.Add();
                finalSDWorksheet.Name = "Final SD Download List";
            }

            if (sourceWorksheet == null)
            {
                MessageBox.Show("Source worksheet not found.");
                return;
            }

            // Get the used range of both source and final worksheets
            Excel.Range sourceRange = sourceWorksheet.UsedRange;
            Excel.Range finalRange = finalSDWorksheet.UsedRange;

            // Iterate over each row in the final worksheet
            foreach (Excel.Range finalRow in finalRange.Rows)
            {
                string finalIssueId = finalRow.Cells[1, 1]?.Value?.ToString(); // Assuming IssueID is in the first column

                if (string.IsNullOrEmpty(finalIssueId))
                {
                    continue; // Skip processing if IssueID is null or empty
                }

                // Iterate over each row in the source worksheet
                foreach (Excel.Range sourceRow in sourceRange.Rows)
                {
                    string sourceIssueId = sourceRow.Cells[1, 1]?.Value?.ToString(); // Assuming IssueID is in the first column

                    if (!string.IsNullOrEmpty(sourceIssueId))
                    {
                        // Check if the final IssueID matches an IssueID in the source sheet or is a variant
                        if (finalIssueId == sourceIssueId || finalIssueId.StartsWith(sourceIssueId + "_"))
                        {
                            // Copy all other column data from the source row to the final row
                            for (int j = 2; j <= sourceWorksheet.Columns.Count; j++) // Start from column 2 to skip the IssueID column
                            {
                                finalRow.Cells[1, j].Value = sourceRow.Cells[1, j].Value;
                            }
                            break; // Stop searching for a match once found
                        }
                    }
                }
            }

            MessageBox.Show("Data copied to 'Final SD Download List' sheet based on matching issue IDs.");
        }

    }
}
