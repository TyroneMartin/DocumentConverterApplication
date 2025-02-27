using System;
using System.IO;
using System.Text;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

public class PdfToHtmlConverter : DocumentConverter
{
    public override string ExpectedInputExtension => ".pdf";
    
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (PDF) to {outputPath} (HTML)...");
        EnsureDirectoryExists(outputPath);

        StringBuilder htmlBuilder = new StringBuilder();
        
        // Add HTML header
        htmlBuilder.AppendLine("<!DOCTYPE html>");
        htmlBuilder.AppendLine("<html>");
        htmlBuilder.AppendLine("<head>");
        htmlBuilder.AppendLine("    <meta charset=\"UTF-8\">");
        htmlBuilder.AppendLine("    <title>Converted PDF Document</title>");
        htmlBuilder.AppendLine("    <style>");
        htmlBuilder.AppendLine("        body { font-family: Arial, sans-serif; line-height: 1.6; }");
        htmlBuilder.AppendLine("    </style>");
        htmlBuilder.AppendLine("</head>");
        htmlBuilder.AppendLine("<body>");

        using (PdfReader reader = new PdfReader(inputPath))
        {
            using (PdfDocument pdfDoc = new PdfDocument(reader))
            {
                for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string text = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(i), strategy);
                    
                    htmlBuilder.AppendLine($"    <div class=\"page\" id=\"page-{i}\">");
                    htmlBuilder.AppendLine($"        <h2>Page {i}</h2>");
                    
                    // Split text by newlines and create paragraphs
                    string[] lines = text.Split(new[] { "\n", "\r\n" }, StringSplitOptions.None);
                    foreach (string line in lines)
                    {
                        if (!string.IsNullOrWhiteSpace(line))
                        {
                            htmlBuilder.AppendLine($"        <p>{line}</p>");
                        }
                    }
                    
                    htmlBuilder.AppendLine("    </div>");
                    htmlBuilder.AppendLine("    <hr/>");
                }
            }
        }
        
        // Close HTML
        htmlBuilder.AppendLine("</body>");
        htmlBuilder.AppendLine("</html>");
        
        // Write to file
        File.WriteAllText(outputPath, htmlBuilder.ToString());

        Console.WriteLine("Conversion complete!");
    }
}
