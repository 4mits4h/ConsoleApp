using ClosedXML.Excel;
using ClosedXML.Excel.Exceptions;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;
using System.Text;
class Program
{
    static void Main(string[] args)
    {
        string excelFilePath = "D:\\Test Projects\\excel\\data.xlsx";
        string outputFile = "D:\\Test Projects\\excel\\UnmatchedData.txt";

        List<string> unmatchedData = new List<string>();

        using (var workbook = new XLWorkbook(excelFilePath))
        {
            var worksheet = workbook.Worksheets.FirstOrDefault(); // Assuming you're working with the first worksheet.

            if (worksheet != null)
            {      
                var columnAData = worksheet.Column("A").Cells().Select(cell => cell.Value.ToString()).ToList();
                var columnBData = worksheet.Column("B").Cells().Select(cell => cell.Value.ToString().Equals(string.Empty)).ToList();
               

                    for (int i = 0; i < Math.Min(columnAData.Count, columnBData.Count); i++)
                {
                    
                    //if (columnAData[i] != columnBData[i])
                    //{
                    //    unmatchedData.Add($"A:{columnAData[i]}, B:{columnBData[i]}");
                    //}
                }

                // Write unmatched data to a text file
                File.WriteAllLines(outputFile, unmatchedData);

                Console.WriteLine("Unmatched data written to " + outputFile);
            }
            else
            {
                Console.WriteLine("Worksheet not found in the Excel file.");
            }///et
        }
    }
}
