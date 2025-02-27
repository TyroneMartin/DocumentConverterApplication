using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

public class PdfToDocxConverter : DocumentConverter
{
    public override string ExpectedInputExtension => ".pdf";
    
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (PDF) to {outputPath} (DOCX)...");
        EnsureDirectoryExists(outputPath);

        // Extract text content from PDF
        StringBuilder textBuilder = new StringBuilder();
        
        using (PdfReader reader = new PdfReader(inputPath))
        {
            using (PdfDocument pdfDoc = new PdfDocument(reader))
            {
                for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string text = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(i), strategy);
                    textBuilder.AppendLine(text);
                    textBuilder.AppendLine(); // Add an extra line between pages
                }
            }
        }
        
        // Create DOCX file
        using (WordprocessingDocument docx = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
        {
            // Add a main document part
            MainDocumentPart mainPart = docx.AddMainDocumentPart();
            
            // Create document structure
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());
            
            // Split text by lines and add as paragraphs
            string[] lines = textBuilder.ToString().Split(
                new[] { "\r\n", "\r", "\n" },
                StringSplitOptions.None
            );
            
            foreach (string line in lines)
            {
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text(line));
            }
            
            mainPart.Document.Save();
        }

        Console.WriteLine("Conversion complete!");
    }
}
