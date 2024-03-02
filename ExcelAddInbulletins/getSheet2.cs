using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInbulletins
{
    internal class getSheet2
    {
        List<string> existIds = new List<string>();
        List<string> issueIds = new List<string>();
        public void CopyDataToUnderscoreSDWorksheet()
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (workbook == null)
            {
                MessageBox.Show("No active workbook found.");
                return;
            }

            Excel.Worksheet sourceWorksheet = workbook.Sheets["Remove_Redundant_Column "];
            Excel.Worksheet underscoreSDWorksheet = null;

            // Check if the underscoreSDWorksheet exists, if not, add it
            try
            {
                underscoreSDWorksheet = workbook.Sheets["Check Underscore & SD download"];
            }
            catch
            {
                // underscoreSDWorksheet doesn't exist, so add it
                underscoreSDWorksheet = (Excel.Worksheet)workbook.Sheets.Add();
                underscoreSDWorksheet.Name = "Check Underscore & SD download";
            }

            if (sourceWorksheet == null)
            {
                MessageBox.Show("Source worksheet not found.");
                return;
            }

            // Define the columns you want to copy (e.g., columns A, B, C)
            int[] columnsToCopy = { 1, 2, 3, 4, 5, 6,7,8 }; // Column indices are 1-based

            Excel.Range sourceRange = sourceWorksheet.UsedRange;
            Excel.Range destinationRange = underscoreSDWorksheet.Cells[1, 1];

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
                    underscoreSDWorksheet.Application.CutCopyMode = Excel.XlCutCopyMode.xlCopy;
                }

                destinationRow++;
            }

            Excel.Range allCells = underscoreSDWorksheet.Cells;
            Excel.Borders borders = allCells.Borders;
            borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            borders.Weight = Excel.XlBorderWeight.xlThin;

            MessageBox.Show("Data copied from Source to 'Check Underscore & SD download' sheet.");

            ExcelReader obj1 = new ExcelReader();
            issueIds = obj1.ReadIssueIds();
            CheckIds obj2 = new CheckIds();
            foreach (string id in issueIds)
            {
                Console.WriteLine(id);
                existIds = obj2.CheckFiles(id);
            }

            //mark the red colour for missing issue ids
             obj2.MarkAbsentData(existIds);
            //copy the sheet to final sd download sheet
            getSheet3 ob = new getSheet3();
            ob.CopyHeaderToFinal_SD_Download_List();

            //add the final issue ids in the sheet
            Addissueids ob1 = new Addissueids();
            ob1.AddIds(existIds);
            //copy the other column data from check underscore and sd download sheet 
            ob1.CopyDataToFinal_SD_Download_List();

            //connect with filezila and downlaod the data from the server by passing the existids list reference

            FileZillaClient obj = new FileZillaClient();
            foreach (string id in existIds)
            {
                obj.Download(id);
                //obj.Downloadvariant(id);
            }

            //get summary from the file 
            getsummaryfromfile read=new getsummaryfromfile();
            read.GetSummary(existIds);
            //foreach (string id in existIds)
            //{
            //    Console.WriteLine(id);
            //}
        }

    }
}
