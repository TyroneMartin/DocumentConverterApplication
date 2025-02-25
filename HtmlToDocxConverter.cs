using Xceed.Words.NET;
using System;
using System.IO;
using HtmlAgilityPack;

public class HtmlToDocxConverter : DocumentConverter
{
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (HTML) to {outputPath} (DOCX)...");
        EnsureDirectoryExists(outputPath);

        var htmlContent = File.ReadAllText(inputPath);
        var htmlDoc = new HtmlDocument();
        htmlDoc.LoadHtml(htmlContent);

        var docx = DocX.Create(outputPath);

        foreach (var node in htmlDoc.DocumentNode.SelectNodes("//text()"))
        {
            if (!string.IsNullOrWhiteSpace(node.InnerText))
            {
                docx.InsertParagraph(node.InnerText);
            }
        }

        docx.Save();

        Console.WriteLine("Conversion complete!");
    }
}