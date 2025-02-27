using System;
using System.IO;
using iText.Kernel.Pdf;
using iText.Layout;

public class HtmlToPdfConverter : DocumentConverter
{
    public override string ExpectedInputExtension => ".html";
    
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (HTML) to {outputPath} (PDF)...");
        EnsureDirectoryExists(outputPath);

        // Read HTML content
        string htmlContent = File.ReadAllText(inputPath);
        
        // Convert to PDF using iText HTML2PDF
        using (FileStream pdfDest = new FileStream(outputPath, FileMode.Create))
        {
            ConverterProperties converterProperties = new ConverterProperties();
            HtmlConverter.ConvertToPdf(htmlContent, pdfDest, converterProperties);
        }

        Console.WriteLine("Conversion complete!");
    }
}
