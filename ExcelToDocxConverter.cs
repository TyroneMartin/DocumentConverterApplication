using Xceed.Words.NET;
using System;
using System.IO;
using NPOI.SS.UserModel; 
using NPOI.XSSF.UserModel;

public class ExcelToDocxConverter : DocumentConverter
{
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (Excel) to {outputPath} (DOCX)...");
        EnsureDirectoryExists(outputPath);

        // Load the Excel file
        using (var fileStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        {
            var workbook = new XSSFWorkbook(fileStream);
            var sheet = workbook.GetSheetAt(0); // Get the first sheet

            var docx = DocX.Create(outputPath);

            // Iterate through rows and cells
            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row != null)
                {
                    var paragraph = docx.InsertParagraph();
                    for (int cellIndex = 0; cellIndex < row.LastCellNum; cellIndex++)
                    {
                        var cell = row.GetCell(cellIndex);
                        if (cell != null)
                        {
                            paragraph.Append(cell.ToString()).Append(" ");
                        }
                    }
                }
            }

            docx.Save();
        }

        Console.WriteLine("Conversion complete!");
    }
}