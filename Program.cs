using System;


class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 3)
        {
            ShowUsage();
            return;
        }

        string converterType = args[0];
        string inputPath = args[1];
        string outputPath = args[2];

        try
        {
            DocumentConverter converter = ConverterFactory.CreateConverter(converterType);
            converter.Convert(inputPath, outputPath);
        }
        catch (ArgumentException ex)
        {
            Console.WriteLine(ex.Message);
            ShowUsage();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }

    static void ShowUsage()
    {
        Console.Clear();
        Console.WriteLine("Document Converter - Usage:");
        Console.WriteLine("  dotnet run <converter> <inputPath> <outputPath>");
        Console.WriteLine("\nAvailable converters:");
        Console.WriteLine("  DOCX converters: docx2pdf, docx2html, docx2txt, docx2excel");
        Console.WriteLine("  PDF converters: pdf2docx, pdf2html, pdf2txt");
        Console.WriteLine("  HTML converters: html2docx, html2pdf");
        Console.WriteLine("  Excel converters: excel2docx, excel2pdf");
        Console.WriteLine("\nExample:");
        Console.WriteLine("  dotnet run docx2pdf document.docx output.pdf");
    }
}
