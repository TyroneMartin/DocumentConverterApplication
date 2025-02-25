using System;
using System.IO;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

public class ExcelToPdfConverter : DocumentConverter
{
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (Excel) to {outputPath} (PDF)...");
        EnsureDirectoryExists(outputPath);

        using (var fileStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        {
            var workbook = new XSSFWorkbook(fileStream);
            var sheet = workbook.GetSheetAt(0);

            using (PdfWriter writer = new PdfWriter(outputPath))
            {
                using (PdfDocument pdfDoc = new PdfDocument(writer))
                {
                    Document document = new Document(pdfDoc);

                    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        var row = sheet.GetRow(rowIndex);
                        if (row != null)
                        {
                            var paragraph = new Paragraph();
                            for (int cellIndex = 0; cellIndex < row.LastCellNum; cellIndex++)
                            {
                                var cell = row.GetCell(cellIndex);
                                if (cell != null)
                                {
                                    paragraph.Add(cell.ToString() + " ");
                                }
                            }
                            document.Add(paragraph);
                        }
                    }

                    document.Close();
                }
            }
        }

        Console.WriteLine("Conversion complete!");
    }
}