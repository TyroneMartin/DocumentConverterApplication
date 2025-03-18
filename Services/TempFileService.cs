using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace DocumentConverterWebApp.Services
{
    public class TempFileService : ITempFileService
    {
        private readonly string _tempDir;
        private readonly TimeSpan _fileExpiryTime = TimeSpan.FromHours(1);

        public TempFileService()
        {
            // Create a temp directory in the application directory
            _tempDir = Path.Combine(Directory.GetCurrentDirectory(), "TempFiles");

            // Ensure the directory exists
            if (!Directory.Exists(_tempDir))
            {
                Directory.CreateDirectory(_tempDir);
            }
        }

        public void CleanupOldFiles()
        {
            CleanupTempFiles(_fileExpiryTime);
        }

        public async Task<(string filePath, string fileName)> SaveUploadedFileAsync(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                throw new ArgumentException("Invalid file");
            }

            var fileName = Path.GetFileName(file.FileName);
            var filePath = Path.Combine(_tempDir, Guid.NewGuid().ToString() + Path.GetExtension(fileName));

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            return (filePath, fileName);
        }

        public string CreateTempFilePath(string extension)
        {
            return Path.Combine(_tempDir, $"{Guid.NewGuid()}{extension}");
        }

        public FileStreamResult GetDownloadFileStream(string filePath, string contentType, string fileName)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("The requested file was not found", filePath);
            }

            var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            return new FileStreamResult(stream, contentType)
            {
                FileDownloadName = fileName
            };
        }

        // Implement ITempFileService methods
        public string GetTempDirectory()
        {
            return _tempDir;
        }

        public void EnsureDirectoryExists(string filePath)
        {
            string? directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
        }

        public void CleanupTempFiles(TimeSpan olderThan)
        {
            try
            {
                var cutoffTime = DateTime.Now - olderThan;
                
                // Clean files in the main temp directory
                if (Directory.Exists(_tempDir))
                {
                    foreach (var file in Directory.GetFiles(_tempDir))
                    {
                        try
                        {
                            var fileInfo = new FileInfo(file);
                            if (fileInfo.CreationTime < cutoffTime)
                            {
                                File.Delete(file);
                            }
                        }
                        catch (Exception ex)
                        {
                            // Log error but continue with other files
                            Console.WriteLine($"Error deleting file {file}: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log the error but don't crash the application
                Console.WriteLine($"Error during temp file cleanup: {ex.Message}");
            }
        }
    }
}