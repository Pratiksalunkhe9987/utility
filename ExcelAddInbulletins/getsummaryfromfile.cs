using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Word.Application;

namespace ExcelAddInbulletins
{
    internal class getsummaryfromfile
    {
        Dictionary<string, string> summary = new Dictionary<string, string>();
        public Dictionary<string, string> GetSummary(List<string> existIds)
        {
            // Assuming the folder path where the summary files are stored
            string folderPath = @"C:\Users\psalunkhe\Desktop\Macro\download";

            // Create an instance of the Word application
            Application wordApp = new Application();

            // Variable to store the message to display in the message box
            string message = "Summary Data:\n";

            foreach (string existId in existIds)
            {
                // Generate the file path based on the existId
                string filePath = Path.Combine(folderPath, existId);

                // Check if the file exists with .docx extension
                if (File.Exists(filePath + ".docx"))
                {
                    filePath += ".docx";
                }
                else if (File.Exists(filePath + ".doc")) // Check if the file exists with .doc extension
                {
                    filePath += ".doc";
                }
                else
                {
                    // Handle case where file does not exist for the existId
                    // For example, set a default summary or skip this existId
                    summary[existId] = "Summary not available";
                    continue;
                }

                // Open the Word document
                Document doc = wordApp.Documents.Open(filePath);

                // Read the content after the "ISSUE SUMMARY:" section
                string fileContents = ReadContentAfterIssueSummary(doc);

                // Close the document without saving changes
                doc.Close(WdSaveOptions.wdDoNotSaveChanges);

                // Add the summary to the dictionary with the existId as the key
                summary[existId] = fileContents;

                // Add the summary to the message to display in the message box
                message += $"Issue ID: {existId},{fileContents}\n";
            }

            // Quit the Word application
            wordApp.Quit();

            // Display the message box with the summary data
            MessageBox.Show(message, "Summary Data", MessageBoxButtons.OK, MessageBoxIcon.Information);

            return summary;
        }

        // Function to read the content after the "ISSUE SUMMARY:" section
        private string ReadContentAfterIssueSummary(Document doc)
        {
            // Initialize the content variable
            string content = "";

            // Find the paragraph containing "ISSUE SUMMARY:"
            foreach (Paragraph para in doc.Paragraphs)
            {
                if (para.Range.Text.Contains("ISSUE SUMMARY:"))
                {
                    // Get the index of the "ISSUE SUMMARY:" paragraph
                    int index = para.Range.Text.IndexOf("ISSUE SUMMARY:") + "ISSUE SUMMARY:".Length;

                    // Read the content after "ISSUE SUMMARY:"
                    content = para.Range.Text.Substring(index).Trim();

                    // Break the loop after finding the "ISSUE SUMMARY:" paragraph
                    break;
                }
            }

            return content;
        }

        /*public Dictionary<string, string> GetSummary(List<string> existIds)
        {
            // Assuming the folder path where the summary files are stored
            string folderPath = @"C:\Users\psalunkhe\Desktop\Macro\download";

            // Create an instance of the Word application
            Application wordApp = new Application();

            // Variable to store the message to display in the message box
            string message = "Summary Data:\n";

            foreach (string existId in existIds)
            {
                // Generate the file path based on the existId
                string filePath = Path.Combine(folderPath, existId);

                // Check if the file exists with .docx extension
                if (File.Exists(filePath + ".docx"))
                {
                    filePath += ".docx";
                }
                else if (File.Exists(filePath + ".doc")) // Check if the file exists with .doc extension
                {
                    filePath += ".doc";
                }
                else
                {
                    // Handle case where file does not exist for the existId
                    // For example, set a default summary or skip this existId
                    summary[existId] = "Summary not available";
                    continue;
                }

                // Open the Word document
                Document doc = wordApp.Documents.Open(filePath);

                // Read the content after the "ISSUE SUMMARY:" section
                string fileContents = ReadContentAfterIssueSummary(doc);

                // Close the document without saving changes
                doc.Close(WdSaveOptions.wdDoNotSaveChanges);

                // Add the summary to the dictionary with the existId as the key
                summary[existId] = fileContents;

                // Add the summary to the message to display in the message box
                message += $"Issue ID: {existId},{fileContents}\n";
            }

            // Quit the Word application
            wordApp.Quit();

            // Display the message box with the summary data
            MessageBox.Show(message, "Summary Data", MessageBoxButtons.OK, MessageBoxIcon.Information);

            return summary;
        }

        // Function to read the content after the "ISSUE SUMMARY:" section
        private string ReadContentAfterIssueSummary(Document doc)
        {
            // Initialize the content variable
            string content = "";

            // Find the paragraph containing "ISSUE SUMMARY:"
            foreach (Paragraph para in doc.Paragraphs)
            {
                if (para.Range.Text.Contains("ISSUE SUMMARY:"))
                {
                    // Get the index of the "ISSUE SUMMARY:" paragraph
                    int index = para.Range.Text.IndexOf("ISSUE SUMMARY:") + "ISSUE SUMMARY:".Length;

                    // Read the content after "ISSUE SUMMARY:"
                    content = para.Range.Text.Substring(index).Trim();

                    // If the content is empty, check for bullet points format
                    if (string.IsNullOrWhiteSpace(content))
                    {
                        // Iterate through the subsequent paragraphs to read bullet points
                        foreach (Paragraph nextPara in doc.Paragraphs)
                        {
                            if (nextPara.Range.Start > para.Range.Start)
                            {
                                // Adjust the condition to match the bullet point format used in your documents
                                if (nextPara.Range.Text.Trim().StartsWith("-") || nextPara.Range.Text.Trim().StartsWith("*") || nextPara.Range.Text.Trim().StartsWith("1."))
                                {
                                    content += nextPara.Range.Text.Trim() + Environment.NewLine;
                                }
                                else
                                {
                                    break; // Exit loop if the bullet point format ends
                                }
                            }
                        }
                    }

                    // Break the loop after finding the "ISSUE SUMMARY:" paragraph
                    break;
                }
            }

            return content;
        }*/

       

    }
}
