using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInbulletins
{
    internal class CheckIds
    {
        private const string FtpServer = "support.e-emphasys.com";
        private const string FtpUsername = "extendftp";
        private const string FtpPassword = "eXte$d@2@!6";
        List<string> matchingFiles = new List<string>();
        /*    public List<string> CheckAndAddIds(List<string> issueIds)
            {
                List<string> existingIds = new List<string>();

                try
                {
                    using (WebClient client = new WebClient())
                    {
                        client.Credentials = new NetworkCredential(FtpUsername, FtpPassword);

                        foreach (string issueId in issueIds)
                        {
                            string remoteDirectory = $"E50C_1_E501/documents/";

                            // Get a list of files in the directory
                            string fileList = client.DownloadString($"ftp://{FtpServer}/{remoteDirectory}");
                            string[] files = fileList.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

                            // Check if any file in the directory contains the issue ID
                            foreach (string file in files)
                            {
                                if (file.Contains(issueId))
                                {
                                    existingIds.Add(issueId);
                                   //break; // Exit the loop after finding the first match
                                }
                            }
                        }
                    }
                }
                catch (WebException ex)
                {
                    Console.WriteLine($"FTP Error: {ex.Message}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                }

                return existingIds;
            }*/
        /*  public List<string> CheckAndAddIds(string issueId)
          {
              List<string> foundFiles = new List<string>();

              try
              {
                  using (WebClient client = new WebClient())
                  {
                      client.Credentials = new NetworkCredential(FtpUsername, FtpPassword);

                      string remoteDirectory = $"E50C_1_E501/documents/";

                      try
                      {
                          // Get a list of files in the directory
                          string fileList = client.DownloadString($"ftp://{FtpServer}/{remoteDirectory}");
                          string[] files = fileList.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

                          // Check if any file in the directory contains the issue ID
                          foreach (string file in files)
                          {
                              if (file.Contains(issueId))
                              {
                                  // Add the file name without extension to the list
                                  string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file);
                                  foundFiles.Add(fileNameWithoutExtension);
                              }
                          }
                      }
                      catch (WebException ex)
                      {
                          Console.WriteLine($"FTP Error: {ex.Message}");
                          // Handle specific FTP-related errors here
                      }
                  }
              }
              catch (Exception ex)
              {
                  Console.WriteLine($"An error occurred: {ex.Message}");
                  // Handle other general exceptions here
              }

              return foundFiles;
          }*/

        public List<string> CheckFiles(string issueId)
        {
            try
            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create($"ftp://{FtpServer}/E50C_1_E501/documents/");
                request.Method = WebRequestMethods.Ftp.ListDirectory;
                request.Credentials = new NetworkCredential(FtpUsername, FtpPassword);

                using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
                using (Stream responseStream = response.GetResponseStream())
                using (StreamReader reader = new StreamReader(responseStream))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (line.Contains(issueId))
                        {
                            // Extract the file name without extension
                            string fileName = Path.GetFileNameWithoutExtension(line);
                            matchingFiles.Add(fileName);
                        }
                    }
                }
                foreach (var item in matchingFiles)
                {
                    Console.WriteLine(item);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            return matchingFiles;
        }

        public void MarkAbsentData(List<string> existIds)
        {
            Excel.Worksheet worksheet = null;

            try
            {
                // Access the current active workbook
                Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (workbook == null)
                {
                    Console.WriteLine("No active workbook found.");
                    return;
                }

                // Access the worksheet where you want to mark the data
                worksheet = workbook.Sheets["Check Underscore & SD download"]; // Replace "Check Underscore & SD download" with the actual worksheet name
                if (worksheet == null)
                {
                    Console.WriteLine("Worksheet not found.");
                    return;
                }

                // Get the used range of the worksheet
                Excel.Range usedRange = worksheet.UsedRange;

                // Get the column index of the issue IDs column (assuming it's the first column)
                int issueIdColumnIndex = 1;

                // Iterate over each row in the used range (start from the second row)
                for (int i = 2; i <= usedRange.Rows.Count; i++)
                {
                    // Get the value of the issue ID column in the current row
                    object cellValue = usedRange.Cells[i, issueIdColumnIndex].Value;
                    string issueId = cellValue != null ? cellValue.ToString() : null;

                    // Check if the issue ID is absent in the existIds list
                    if (issueId != null && !existIds.Contains(issueId))
                    {
                        // Set the interior color of the entire row to red
                        Excel.Range row = (Excel.Range)worksheet.Rows[i];
                        row.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    }
                }

                Console.WriteLine("Data marked for absent issue IDs.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Release COM objects
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }



    }

}
