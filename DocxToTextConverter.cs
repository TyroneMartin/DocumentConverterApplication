using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class DocxToTextConverter : DocumentConverter
{
    public override string ExpectedInputExtension => ".docx";
    
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (DOCX) to {outputPath} (TXT)...");
        EnsureDirectoryExists(outputPath);

        StringBuilder txtBuilder = new StringBuilder();
        
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
                        txtBuilder.AppendLine(text);
                    }
                }
            }
        }
        
        // Write to file
        File.WriteAllText(outputPath, txtBuilder.ToString());

        Console.WriteLine("Conversion complete!");
    }
}
