using System;
using System.IO;

// Base abstract class for all converters
public abstract class DocumentConverter
{
    public abstract string ExpectedInputExtension { get; }
    
    public void ConvertWithValidation(string inputPath, string outputPath)
    {
        try
        {
            // Check if input file exists
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException($"Input file not found: {inputPath}");
            }
            
            // Validate file extension
            string extension = Path.GetExtension(inputPath).ToLowerInvariant();
            if (!extension.Equals(ExpectedInputExtension, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException(
                    $"Invalid input file format. Expected {ExpectedInputExtension} but got {extension}. " + 
                    $"This converter only works with {ExpectedInputExtension.TrimStart('.')} files."
                );
            }
            
            // Call the actual conversion implementation
            Convert(inputPath, outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
            throw;
        }
    }
    
    // Abstract method to be implemented by derived classes
    public abstract void Convert(string inputPath, string outputPath);

    protected void EnsureDirectoryExists(string filePath)
    {
        // Use nullable string to handle potential null values
        string? directory = Path.GetDirectoryName(filePath);

        // Check if the directory is null or empty b4 creating it
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }
}