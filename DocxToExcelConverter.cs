using Xceed.Words.NET;
using System;
using System.IO;
using NPOI.XSSF.UserModel;

public class DocxToExcelConverter : DocumentConverter
{
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (DOCX) to {outputPath} (Excel)...");
        EnsureDirectoryExists(outputPath);

        var docx = DocX.Load(inputPath);
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Converted Document");

        int rowIndex = 0;
        foreach (var paragraph in docx.Paragraphs)
        {
            if (!string.IsNullOrWhiteSpace(paragraph.Text))
            {
                var row = sheet.CreateRow(rowIndex++);
                var cell = row.CreateCell(0);
                cell.SetCellValue(paragraph.Text);
            }
        }

        using (var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        {
            workbook.Write(fileStream);
        }

        Console.WriteLine("Conversion complete!");
    }
}