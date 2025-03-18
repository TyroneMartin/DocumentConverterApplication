using System;
using System.IO;
using ConverterLibrary;

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
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' does not exist");
                return;
            }

            DocumentConverter converter = ConverterFactory.CreateConverter(converterType);
            
            converter.ConvertWithValidation(inputPath, outputPath);
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
        Console.WriteLine("Welcome to the Document Converter Application!\n");
        Console.WriteLine("Document Converter - Usage:");
        Console.WriteLine("  In Terminal Type:  dotnet run <converter> <inputPath> <outputPath>");
        Console.WriteLine("\nAvailable converters:");
        Console.WriteLine("  DOCX converters: docx2pdf, docx2html, docx2txt, docx2excel");
        Console.WriteLine("  PDF converters: pdf2docx, pdf2txt"); // pdf2html not yet implemented
        Console.WriteLine("  HTML converters: html2docx"); // html2pdf not yet implemented
        // Console.WriteLine("  Excel converters: excel2docx, excel2pdf");
        Console.WriteLine("\nExamples:");
        Console.WriteLine("  Example: dotnet run docx2pdf ./Documents/Doc6.docx ./Documents/output.pdf" );
        Console.WriteLine("  dotnet run pdf2txt document.pdf output.txt");
    }
}