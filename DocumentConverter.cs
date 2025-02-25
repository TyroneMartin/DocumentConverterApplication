using System;
using System.IO;

// Base abstract class for all converters
public abstract class DocumentConverter
{
    public abstract void Convert(string inputPath, string outputPath);

    protected void EnsureDirectoryExists(string filePath)
    {
        // Use nullable string to handle potential null values
        string? directory = Path.GetDirectoryName(filePath);

        // Check if the directory is not null or empty before creating it
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
    }
}