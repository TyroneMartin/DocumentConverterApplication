using Xceed.Words.NET;
using System;
using System.IO;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;

public class DocxToPdfConverter : DocumentConverter
{
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (DOCX) to {outputPath} (PDF)...");
        EnsureDirectoryExists(outputPath);

        var docx = DocX.Load(inputPath);

        using (PdfWriter writer = new PdfWriter(outputPath))
        {
            using (PdfDocument pdfDoc = new PdfDocument(writer))
            {
                Document document = new Document(pdfDoc);

                foreach (var paragraph in docx.Paragraphs)
                {
                    document.Add(new Paragraph(paragraph.Text));
                }

                document.Close();
            }
        }

        Console.WriteLine("Conversion complete!");
    }
}