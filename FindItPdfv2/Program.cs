using ClosedXML.Excel;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

class Program
{
    static void Main(string[] args)
    {
        // Path to your folder
        string folderPath = @"C:\Users\User\source\repos\FindItPdfv2\FindItPdfv2\TestFolder\";

        // Keyword to search for
        string keyword = "einziehen.";

        if (Directory.Exists(folderPath))
        {
            // Get all PDF files in the specified folder
            string[] pdfFiles = Directory.GetFiles(folderPath, "*.pdf");

            // Create a new Excel workbook
            using (var workbook = new XLWorkbook())
            {
                // Add a worksheet to the workbook
                var worksheet = workbook.Worksheets.Add("Keyword Results");

                // Set the headers for the Excel sheet
                worksheet.Cell(1, 1).Value = "File Name";
                worksheet.Cell(1, 2).Value = "Keyword Values";
                worksheet.Cell(1, 3).Value = "Address";

                int currentRow = 2; // Start from the second row

                // Process each PDF file
                foreach (var pdfPath in pdfFiles)
                {
                    try
                    {
                        // Extract text from the PDF file
                        string allTextInPdf = ExtractTextFromPdf(pdfPath);

                        // Find values associated with the keyword
                        var keywordValue = FindKeywordValues(allTextInPdf, keyword);

                        // Find addresses in the text
                        var addresses = FindAddresses(allTextInPdf);

                        // Write file name in the first column
                        worksheet.Cell(currentRow, 1).Value = Path.GetFileName(pdfPath);

                        // Write keyword results in the second column
                        if (keywordValue.ContainsKey(keyword))
                        {
                            worksheet.Cell(currentRow, 2).Value = string.Join(", ", keywordValue[keyword]);
                        }
                        else
                        {
                            worksheet.Cell(currentRow, 2).Value = "Not Found";
                        }

                        // Write the address in the third column
                        worksheet.Cell(currentRow, 3).Value = addresses;

                        currentRow++; // Move to the next row
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file {pdfPath}: {ex.Message}");
                    }
                }

                // Define the path where you want to save the Excel file
                string outputFilePath = @"C:\Users\User\source\repos\FindItPdfv2\FindItPdfv2\NewFolder\Table2.xlsx";

                // Save the workbook to the specified path
                workbook.SaveAs(outputFilePath);
                Console.WriteLine($"Excel file '{outputFilePath}' has been created.");
            }
        }
        else
        {
            Console.WriteLine($"The directory '{folderPath}' does not exist.");
        }
    }

    // Extract text from PDF using iText
    static string ExtractTextFromPdf(string pdfPath)
    {
        using (PdfReader reader = new PdfReader(pdfPath))
        using (PdfDocument pdfDoc = new PdfDocument(reader))
        {
            StringBuilder text = new StringBuilder();

            // Iterate through each page of the PDF
            for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
            {
                ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();

                // Extract text from the page
                string pageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                text.Append(pageText);
            }

            return text.ToString();
        }
    }

    // Find keyword values using regex
    static Dictionary<string, List<string>> FindKeywordValues(string text, string keyword)
    {
        var results = new Dictionary<string, List<string>>();

        // Regex pattern to capture the number after "einziehen"
        // Adjust pattern based on how the number is formatted in your PDF
        string pattern = $@"{Regex.Escape(keyword)}\s+(\d{{1,3}}(?:\.\d{{3}})*,\d{{2}})";
        MatchCollection matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);


        foreach (Match match in matches)
        {
            if (!results.ContainsKey(keyword))
            {
                results[keyword] = new List<string>();
            }

            // Add the number after "einziehen" to the results
            results[keyword].Add(match.Groups[1].Value.Trim());
        }

        return results;
    }

    // Find addresses with hardcoded values
    static string FindAddresses(string text)
    {
        // Hardcoded address patterns
        var hardcodedAddresses = new List<string>
        {
            "D - 12099 Berlin",
            "D - 21031 Hamburg",
            "D - 24857 Fahrdorf",
            "D - 37242 Bad Sooden-Allendorf",
            "D - 15344 Strausberg",
            "D - 38644 Goslar"
        };

        // Regular expression pattern for matching hardcoded addresses
        string pattern = string.Join("|", hardcodedAddresses.Select(Regex.Escape));

        // Find matches
        MatchCollection matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);

        // Return the first matched address or default if none found
        if (matches.Count > 0)
        {
            return matches[0].Value;
        }
        else
        {
            // Default address if none of the hardcoded values are found
            return "Default Address: D - 38165 Lehre";
        }
    }
}
