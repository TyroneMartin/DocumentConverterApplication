using System;
using System.IO;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class DocxToExcelConverter : DocumentConverter
{
    public override string ExpectedInputExtension => ".docx";
    
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (DOCX) to {outputPath} (Excel)...");
        EnsureDirectoryExists(outputPath);

        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Converted Document");

        // Open the document using Open XML SDK
        using (WordprocessingDocument doc = WordprocessingDocument.Open(inputPath, false))
        {
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body != null)
            {
                int rowIndex = 0;

                // Extract paragraphs and add to Excel
                foreach (var paragraph in body.Descendants<Paragraph>())
                {
                    string text = paragraph.InnerText;
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        var row = sheet.CreateRow(rowIndex++);
                        var cell = row.CreateCell(0);
                        cell.SetCellValue(text);
                    }
                }
            }
        }

        using (var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        {
            workbook.Write(fileStream);
        }

        Console.WriteLine("Conversion complete!");
    }
}
