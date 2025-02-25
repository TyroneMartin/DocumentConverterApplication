using System;
using System.IO;
using HtmlAgilityPack;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;

public class HtmlToPdfConverter : DocumentConverter
{
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (HTML) to {outputPath} (PDF)...");
        EnsureDirectoryExists(outputPath);

        var htmlDoc = new HtmlDocument();
        htmlDoc.Load(inputPath);

        using (PdfWriter writer = new PdfWriter(outputPath))
        {
            using (PdfDocument pdfDoc = new PdfDocument(writer))
            {
                Document document = new Document(pdfDoc);

                var textNodes = htmlDoc.DocumentNode.SelectNodes("//text()");
                if (textNodes != null)
                {
                    foreach (var node in textNodes)
                    {
                        string text = node.InnerText.Trim();
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            document.Add(new Paragraph(text));
                        }
                    }
                }

                document.Close();
            }
        }

        Console.WriteLine("Conversion complete!");
    }
}