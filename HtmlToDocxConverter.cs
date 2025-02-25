using System;
using System.IO;
using HtmlAgilityPack;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class HtmlToDocxConverter : DocumentConverter
{
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (HTML) to {outputPath} (DOCX)...");
        EnsureDirectoryExists(outputPath);

        // Load HTML document
        var htmlDoc = new HtmlDocument();
        htmlDoc.Load(inputPath);

        // Create Word document
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputPath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            // Add a main document part
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
            
            // Create the document structure
            mainPart.Document = new Document();
            Body body = new Body();
            mainPart.Document.Append(body);

            // Extract text from HTML and add to Word document
            var textNodes = htmlDoc.DocumentNode.SelectNodes("//text()[normalize-space()]");
            if (textNodes != null)
            {
                foreach (var textNode in textNodes)
                {
                    string text = textNode.InnerText.Trim();
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        Paragraph para = new Paragraph();
                        Run run = new Run();
                        Text textElement = new Text(text);
                        
                        run.Append(textElement);
                        para.Append(run);
                        body.Append(para);
                    }
                }
            }
        }

        Console.WriteLine("Conversion complete!");
    }
}