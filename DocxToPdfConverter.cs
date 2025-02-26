using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;

public class DocxToPdfConverter : DocumentConverter
{
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (DOCX) to {outputPath} (PDF)...");
        EnsureDirectoryExists(outputPath);

        // Create PDF document using iText7
        using (PdfWriter writer = new PdfWriter(outputPath))
        {
            using (PdfDocument pdfDoc = new PdfDocument(writer))
            {
                iText.Layout.Document document = new iText.Layout.Document(pdfDoc);

                // Extract content from DOCX
                using (WordprocessingDocument doc = WordprocessingDocument.Open(inputPath, false))
                {
                    if (doc.MainDocumentPart?.Document?.Body != null)
                    {
                        var body = doc.MainDocumentPart.Document.Body;

                        // Process paragraphs
                        foreach (var paragraph in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                        {
                            string text = paragraph.InnerText;
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                document.Add(new iText.Layout.Element.Paragraph(text));
                            }
                        }
                    }
                }

                document.Close();
            }
        }

        Console.WriteLine("Conversion complete!");
    }
}