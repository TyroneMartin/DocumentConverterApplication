using System;
using System.IO;
using iText.Kernel.Pdf; 
using iText.Kernel.Pdf.Canvas.Parser; 
using iText.Kernel.Pdf.Canvas.Parser.Listener; 
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class PdfToDocxConverter : DocumentConverter
{
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (PDF) to {outputPath} (DOCX)...");
        EnsureDirectoryExists(outputPath);

        // Create Word document
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputPath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            // Add a main document part
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = new Body();
            mainPart.Document.Append(body);

            // Extract text from PDF using iText7
            using (PdfReader reader = new PdfReader(inputPath))
            {
                using (PdfDocument pdfDoc = new PdfDocument(reader))
                {
                    for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                    {
                        ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                        string text = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(i), strategy);

                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            // Add extracted text to Word document
                            Paragraph para = new Paragraph();
                            Run run = new Run();
                            Text txt = new Text(text);
                            run.Append(txt);
                            para.Append(run);
                            body.Append(para);
                        }
                    }
                }
            }
        }

        Console.WriteLine("Conversion complete!");
    }
}