using System;
using System.IO;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;

public class PdfToTextConverter : DocumentConverter
{
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (PDF) to {outputPath} (TXT)...");
        EnsureDirectoryExists(outputPath);

        using (PdfReader reader = new PdfReader(inputPath))
        {
            using (PdfDocument pdfDoc = new PdfDocument(reader))
            {
                using (StreamWriter writer = new StreamWriter(outputPath))
                {
                    for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                    {
                        var page = pdfDoc.GetPage(i);
                        var text = PdfTextExtractor.GetTextFromPage(page);
                        writer.WriteLine(text);
                    }
                }
            }
        }

        Console.WriteLine("Conversion complete!");
    }
}