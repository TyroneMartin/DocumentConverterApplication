using Xceed.Words.NET;
using System;
using System.IO;
using System.Text;

public class DocxToTextConverter : DocumentConverter
{
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (DOCX) to {outputPath} (TXT)...");
        EnsureDirectoryExists(outputPath);

        var docx = DocX.Load(inputPath);
        
        StringBuilder text = new StringBuilder();
        foreach (var paragraph in docx.Paragraphs)
        {
            if (!string.IsNullOrWhiteSpace(paragraph.Text))
            {
                text.AppendLine(paragraph.Text);
            }
        }

        File.WriteAllText(outputPath, text.ToString());

        Console.WriteLine("Conversion complete!");
    }
}