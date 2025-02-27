using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Kernel.Pdf;
using iText.Layout;

public class DocxToPdfConverter : DocumentConverter
{
    public override string ExpectedInputExtension => ".docx";
    
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (DOCX) to {outputPath} (PDF)...");
        
        try
        {
            EnsureDirectoryExists(outputPath);
            
            StringBuilder contentBuilder = new StringBuilder();
            
            using (WordprocessingDocument doc = WordprocessingDocument.Open(inputPath, false))
            {
                if (doc.MainDocumentPart?.Document?.Body == null)
                {
                    throw new InvalidOperationException("The DOCX document appears to be empty or corrupted.");
                }
                
                var body = doc.MainDocumentPart.Document.Body;
                
                // Process paragraphs
                foreach (var paragraph in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                {
                    string text = paragraph.InnerText;
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        contentBuilder.AppendLine(text);
                    }
                }
            }
            
            using (PdfWriter writer = new PdfWriter(outputPath))
            {
                using (PdfDocument pdfDoc = new PdfDocument(writer))
                {
                    using (iText.Layout.Document document = new iText.Layout.Document(pdfDoc))
                    {
                        // Split the content by lines and add each as a paragraph
                        string[] lines = contentBuilder.ToString().Split(
                            new[] { "\r\n", "\r", "\n" }, 
                            StringSplitOptions.None
                        );
                        
                        foreach (string line in lines)
                        {
                            document.Add(new iText.Layout.Element.Paragraph(line));
                        }
                    }
                }
            }
            
            Console.WriteLine("Conversion complete!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during PDF conversion: {ex.GetType().Name} - {ex.Message}");
            
            if (ex.InnerException != null)
            {
                Console.WriteLine($"Inner exception: {ex.InnerException.GetType().Name} - {ex.InnerException.Message}");
            }
            
            throw;
        }
    }
}