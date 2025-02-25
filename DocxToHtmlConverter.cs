using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class DocxToHtmlConverter : DocumentConverter
{
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (DOCX) to {outputPath} (HTML)...");
        EnsureDirectoryExists(outputPath);

        StringBuilder htmlBuilder = new StringBuilder();
        
        // Add HTML header
        htmlBuilder.AppendLine("<!DOCTYPE html>");
        htmlBuilder.AppendLine("<html>");
        htmlBuilder.AppendLine("<head>");
        htmlBuilder.AppendLine("    <meta charset=\"UTF-8\">");
        htmlBuilder.AppendLine("    <title>Converted Document</title>");
        htmlBuilder.AppendLine("    <style>");
        htmlBuilder.AppendLine("        body { font-family: Arial, sans-serif; line-height: 1.6; }");
        htmlBuilder.AppendLine("    </style>");
        htmlBuilder.AppendLine("</head>");
        htmlBuilder.AppendLine("<body>");
        
        // Open the document
        using (WordprocessingDocument doc = WordprocessingDocument.Open(inputPath, false))
        {
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body != null)
            {
                // Process paragraphs
                foreach (var paragraph in body.Descendants<Paragraph>())
                {
                    string text = paragraph.InnerText;
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        htmlBuilder.AppendLine($"    <p>{text}</p>");
                    }
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