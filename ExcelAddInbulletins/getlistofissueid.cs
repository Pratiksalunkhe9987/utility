using System.Collections.Generic;
using System;
using Excel = Microsoft.Office.Interop.Excel;

public class ExcelReader
{
    public List<string> ReadIssueIds()
    {
        List<string> issueIds = new List<string>();

        // Hardcoded file path, worksheet name, and column name
        string filePath = @"C:\Users\psalunkhe\Desktop\Bulletinstest.xlsx";
        string worksheetName = "Initial"; // Change to your worksheet name
        string columnName = "IssueID"; // Change to your column name

        // Create an instance of Excel application
        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook workbook = null;

        try
        {
            // Open the Excel workbook
            workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[worksheetName];

            // Find the column index of the specified column name
            Excel.Range columnHeader = worksheet.Rows[1].Find(columnName, Type.Missing,
                Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole,
                Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, false,
                Type.Missing, Type.Missing);

            if (columnHeader != null)
            {
                int columnIndex = columnHeader.Column;

                // Get the used range of the column
                Excel.Range columnRange = worksheet.Columns[columnIndex];
                Excel.Range usedRange = columnRange.Cells.SpecialCells(Excel.XlCellType.xlCellTypeConstants);

                // Read values from the column into an array
                object[,] values = (object[,])usedRange.Value;

                // Iterate over the values and add them to the list
                int rowCount = usedRange.Rows.Count;
                for (int i = 2; i <= rowCount; i++)
                {
                    string value = Convert.ToString(values[i, 1]);
                    issueIds.Add(value);
                }
            }
            else
            {
                Console.WriteLine("Column not found.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
        finally
        {
            // Close Excel workbook and quit application
            workbook?.Close(false);
            excelApp?.Quit();
        }
        return issueIds;
    }
}
