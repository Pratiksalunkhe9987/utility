using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

namespace ExcelAddInbulletins
{
    public partial class ThisAddIn
    {
        public string[] IssueIds { get; private set; }
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Call your method to open the test file when the add-in starts up
            OpenTestFile(@"C:\Users\psalunkhe\Desktop\Bulletinstest.xlsx");
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Clean up any resources if needed
        }

        // Method to open the test file
        private void OpenTestFile(string filePath)
        {
            try
            {
                // Check if the file exists
                if (File.Exists(filePath))
                {
                    // Open the existing workbook
                    Excel.Application excelApp = this.Application;
                    excelApp.Visible = true;
                    Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);

                    // Set number format to "General" for all cells in all worksheets
                    foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                    {
                        Excel.Range usedRange = worksheet.UsedRange;
                        if (usedRange != null)
                        {
                            usedRange.NumberFormat = "General";
                            // Store issue ID in an array from column 1 (A) if the used range is not null
                          //    object[,] data = usedRange.Value;
                          //  int rowCount = data.GetLength(0);
                          // Console.WriteLine(rowCount);
                          //  string[] issueIds = new string[rowCount];
                          //  for (int i = 0; i < rowCount; i++)
                          //  {
                          //      issueIds[i] = data[i + 1, 1]?.ToString(); // Assuming issue ID is in column 1 (A)
                          //  }
                          //  Console.WriteLine(issueIds);
                          //  Console.WriteLine("data copy successfully");
                            // Do further processing with issueIds array as needed
                        }
                    }
                }
                else
                {
                    // Display an error message if the file does not exist
                    System.Windows.Forms.MessageBox.Show("File not found: " + filePath);
                }
            }
            catch (Exception ex)
            {
                // Display an error message if an exception occurs
                System.Windows.Forms.MessageBox.Show("Error opening file: " + ex.Message);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
