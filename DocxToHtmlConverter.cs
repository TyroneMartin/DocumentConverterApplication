using Xceed.Words.NET;
using System;
using System.IO;
using System.Text;
using HtmlAgilityPack;

public class DocxToHtmlConverter : DocumentConverter
{
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (DOCX) to {outputPath} (HTML)...");
        EnsureDirectoryExists(outputPath);

        var docx = DocX.Load(inputPath);
        StringBuilder html = new StringBuilder(); // Fix: Use StringBuilder

        html.AppendLine("<!DOCTYPE html>");
        html.AppendLine("<html>");
        html.AppendLine("<head>");
        html.AppendLine("  <title>Converted Document</title>");
        html.AppendLine("  <style>");
        html.AppendLine("    body { font-family: Arial, sans-serif; }");
        html.AppendLine("  </style>");
        html.AppendLine("</head>");
        html.AppendLine("<body>");

        foreach (var paragraph in docx.Paragraphs)
        {
            if (!string.IsNullOrWhiteSpace(paragraph.Text))
            {
                html.AppendLine($"  <p>{paragraph.Text}</p>");
            }
        }

        html.AppendLine("</body>");
        html.AppendLine("</html>");

        File.WriteAllText(outputPath, html.ToString());

        Console.WriteLine("Conversion complete!");
    }
}