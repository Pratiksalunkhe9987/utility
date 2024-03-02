/*using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInbulletins
{
    internal class FileZillaClient
    {
        private const string FtpServer = "support.e-emphasys.com";
        private const string FtpUsername = "extendftp";
        private const string FtpPassword = "eXte$d@2@!6";


        public void DownloadFile()
        {
            string remoteFilePath = "E50C_1_E501/documents/100010.doc"; // Corrected remote file path
            string localDirectory = @"C:\Users\psalunkhe\Desktop\Macro\download";

            try
            {
                using (WebClient client = new WebClient())
                {
                    client.Credentials = new NetworkCredential(FtpUsername, FtpPassword);
                    client.DownloadFile($"ftp://{FtpServer}/{remoteFilePath}", Path.Combine(localDirectory, "100010.doc"));
                    Console.WriteLine("File downloaded successfully.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

    }
}*/

using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;


namespace ExcelAddInbulletins
{
    internal class FileZillaClient
    {
        private const string FtpServer = "support.e-emphasys.com";
        private const string FtpUsername = "extendftp";
        private const string FtpPassword = "eXte$d@2@!6";

        public void DownloadFile(List<string> issueIds)//string[] issueIds
        {
            string localDirectory = @"C:\Users\psalunkhe\Desktop\Macro\download";

            try
            {
                /* using (WebClient client = new WebClient())
                 {
                     client.Credentials = new NetworkCredential(FtpUsername, FtpPassword);

                     foreach (string issueId in issueIds)
                     {
                         string remoteFilePath = $"E50C_1_E501/documents/{issueId}.doc";
                         string localFilePath = Path.Combine(localDirectory, $"{issueId}.doc");

                         client.DownloadFile($"ftp://{FtpServer}/{remoteFilePath}", localFilePath);
                         Console.WriteLine($"File '{issueId}.doc' downloaded successfully.");
                     }
                 }*/
                using (WebClient client = new WebClient())
                {
                    client.Credentials = new NetworkCredential(FtpUsername, FtpPassword);

                    foreach (string issueId in issueIds)
                    {
                        string remoteDocFilePath = $"E50C_1_E501/documents/{issueId}.doc";
                        string localDocFilePath = Path.Combine(localDirectory, $"{issueId}.doc");

                        string remoteDocxFilePath = $"E50C_1_E501/documents/{issueId}.docx";
                        string localDocxFilePath = Path.Combine(localDirectory, $"{issueId}.docx");

                        // Check if .docx file exists, if not, download .doc file
                        if (File.Exists(localDocxFilePath))
                        {
                            client.DownloadFile($"ftp://{FtpServer}/{remoteDocxFilePath}", localDocxFilePath);
                            Console.WriteLine($"File '{issueId}.docx' downloaded successfully.");
                        }
                        else if (File.Exists(localDocFilePath))
                        {
                            client.DownloadFile($"ftp://{FtpServer}/{remoteDocFilePath}", localDocFilePath);
                            Console.WriteLine($"File '{issueId}.doc' downloaded successfully.");
                        }
                        else
                        {
                            Console.WriteLine($"File '{issueId}' not found in both .doc and .docx formats.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }


        public void Download(string issueId)
        {
            string localDirectory = @"C:\Users\psalunkhe\Desktop\Macro\download";

            try
            {
                using (WebClient client = new WebClient())
                {
                    client.Credentials = new NetworkCredential(FtpUsername, FtpPassword);

                    string remoteFilePathDocx = $"E50C_1_E501/documents/{issueId}.docx"; // Change extension to .docx
                    string localFilePathDocx = Path.Combine(localDirectory, $"{issueId}.docx"); // Change extension to .docx

                    try
                    {
                        // Try downloading in .docx format
                        client.DownloadFile($"ftp://{FtpServer}/{remoteFilePathDocx}", localFilePathDocx);
                        Console.WriteLine($"File '{issueId}.docx' downloaded successfully.");
                    }
                    catch (WebException ex)
                    {
                        if (((FtpWebResponse)ex.Response).StatusCode == FtpStatusCode.ActionNotTakenFileUnavailable)
                        {
                            // If .docx format is unavailable, try downloading in .doc format
                            string remoteFilePathDoc = $"E50C_1_E501/documents/{issueId}.doc"; // Change extension to .doc
                            string localFilePathDoc = Path.Combine(localDirectory, $"{issueId}.doc"); // Change extension to .doc

                            client.DownloadFile($"ftp://{FtpServer}/{remoteFilePathDoc}", localFilePathDoc);
                            Console.WriteLine($"File '{issueId}.doc' downloaded successfully.");
                        }
                        else
                        {
                            throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        /*   public void Downloadvariant(string issueId)
           {
               string localDirectory = @"C:\Users\psalunkhe\Desktop\Macro\download";
               string[] extensionsToTry = { ".docx", ".doc" }; // List of extensions to try

               try
               {
                   using (WebClient client = new WebClient())
                   {
                       client.Credentials = new NetworkCredential(FtpUsername, FtpPassword);

                       foreach (string extension in extensionsToTry)
                       {
                           string remoteDirectory = $"E50C_1_E501/documents/";
                           string[] files = client.DownloadString($"ftp://{FtpServer}/{remoteDirectory}")
                                                 .Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                           foreach (string file in files)
                           {
                               if (file.Contains(issueId) && file.EndsWith(extension)) // Check if the file ends with the current extension
                               {
                                   // Found a file containing the issue ID in its name and with the correct extension, download it
                                   string remoteFilePath = remoteDirectory + file;
                                   string localFilePath = Path.Combine(localDirectory, file);
                                   client.DownloadFile($"ftp://{FtpServer}/{remoteFilePath}", localFilePath);
                                   Console.WriteLine($"File '{file}' downloaded successfully.");
                               }
                           }
                       }
                   }
               }
               catch (Exception ex)
               {
                   Console.WriteLine($"An error occurred: {ex.Message}");
               }
           }*/
       /* public void Downloadvariant(string issueId)
        {
            string localDirectory = @"C:\Users\psalunkhe\Desktop\Macro\download";

            try
            {
                using (WebClient client = new WebClient())
                {
                    client.Credentials = new NetworkCredential(FtpUsername, FtpPassword);

                    string[] extensionsToTry = { ".docx", ".doc" }; // List of extensions to try

                    foreach (string extension in extensionsToTry)
                    {
                        string remoteFilePath = $"E50C_1_E501/documents/{issueId}{extension}";
                        string localFilePath = Path.Combine(localDirectory, $"{issueId}{extension}");

                        try
                        {
                            client.DownloadFile($"ftp://{FtpServer}/{remoteFilePath}", localFilePath);
                            Console.WriteLine($"File '{issueId}{extension}' downloaded successfully.");
                            // If download successful, break out of the loop
                            break;
                        }
                        catch (WebException ex)
                        {
                            if (((FtpWebResponse)ex.Response).StatusCode == FtpStatusCode.ActionNotTakenFileUnavailable)
                            {
                                // If the file is not available, continue to the next extension
                                continue;
                            }
                            else
                            {
                                throw; // Re-throw other exceptions
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }*/
      

        public void Downloadvariant(string issueId)
        {
                 string localDirectory = @"C:\Users\psalunkhe\Desktop\Macro\download";

              try
             {
                  using (WebClient client = new WebClient())
                  {
                     client.Credentials = new NetworkCredential(FtpUsername, FtpPassword);

                    string remoteDirectory = $"E50C_1_E501/documents/";
                    string[] files = client.DownloadString($"ftp://{FtpServer}/{remoteDirectory}")
                                         .Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                // Define a regular expression pattern to match the issue ID within the file names
                string pattern = $@"\b{issueId}\b";

                foreach (string file in files)
                {
                    if (Regex.IsMatch(file, pattern))
                    {
                        // Found a file containing the issue ID in its name, download it
                        string remoteFilePath = remoteDirectory + file;
                        string localFilePath = Path.Combine(localDirectory, file);
                        client.DownloadFile($"ftp://{FtpServer}/{remoteFilePath}", localFilePath);
                        Console.WriteLine($"File '{file}' downloaded successfully.");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
}
