using Xceed.Words.NET;
using System;
using System.IO;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;

public class PdfToDocxConverter : DocumentConverter
{
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (PDF) to {outputPath} (DOCX)...");
        EnsureDirectoryExists(outputPath);

        using (PdfReader reader = new PdfReader(inputPath))
        {
            using (PdfDocument pdfDoc = new PdfDocument(reader))
            {
                var docx = DocX.Create(outputPath);

                for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                {
                    var page = pdfDoc.GetPage(i);
                    var text = PdfTextExtractor.GetTextFromPage(page); // Fix: Use PdfTextExtractor
                    docx.InsertParagraph(text);
                }

                docx.Save();
            }
        }

        Console.WriteLine("Conversion complete!");
    }
}