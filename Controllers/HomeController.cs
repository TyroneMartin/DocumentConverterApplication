using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.IO;
using System;
using DocumentConverterApplication.Models;
using DocumentConverterApplication.Services;

namespace DocumentConverterApplication.Controllers
{
    public class HomeController : Controller
    {
        // Declare a private field for the temp file service.
        private readonly ITempFileService _tempFileService;

        // Constructor injection for the temp file service.
        public HomeController(ITempFileService tempFileService)
        {
            _tempFileService = tempFileService;
        }

        public IActionResult Index()
        {
            var model = new ConverterViewModel
            {
                AvailableConverters = new Dictionary<string, List<string>>
                {
                    { "Text", new List<string> { "txt to pdf", "txt to docx" } },
                    { "Document", new List<string> { "docx to pdf", "pdf to docx" } },
                    { "Spreadsheet", new List<string> { "excel to pdf", "csv to xlsx" } }
                }
            };
            return View(model);
        }

        [HttpPost]
        public IActionResult ConvertDocument(ConverterViewModel model)
        {

            Console.WriteLine("ConvertDocument action invoked.");

            if (!ModelState.IsValid || model.UploadedFile == null || string.IsNullOrEmpty(model.SelectedConverter))
            {
                Console.WriteLine("Model state invalid or missing file/selection.");
                model.IsModelStateValid = false;
                return View("Index", model);
            }

            Console.WriteLine($"Selected converter: {model.SelectedConverter}");
            

            var tempDir = _tempFileService.GetTempDirectory();
            var inputFileName = Path.GetFileName(model.UploadedFile.FileName);
            var tempInputPath = Path.Combine(tempDir, Guid.NewGuid() + Path.GetExtension(inputFileName));

            Console.WriteLine($"Saving uploaded file to temporary path: {tempInputPath}");
            using (var stream = new FileStream(tempInputPath, FileMode.Create))
            {
                model.UploadedFile.CopyTo(stream);
            }

            var outputFileName = "converted_" + inputFileName;
            var downloadsDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "downloads");

            // Ensure the downloads directory exists
            if (!Directory.Exists(downloadsDir))
            {
                Console.WriteLine($"Creating downloads directory: {downloadsDir}");
                Directory.CreateDirectory(downloadsDir);
            }

            var outputPath = Path.Combine(downloadsDir, outputFileName);
            Console.WriteLine($"Absolute output path: {Path.GetFullPath(outputPath)}");

            try
            {
                var converter = ConverterLibrary.ConverterFactory.CreateConverter(
                    model.SelectedConverter.Replace(" ", "").ToLower()
                );
                Console.WriteLine($"Converter selected: {converter.GetType().Name}");
                converter.ConvertWithValidation(tempInputPath, outputPath);

                // Verify the file was created
                if (!System.IO.File.Exists(outputPath))
                {
                    throw new FileNotFoundException($"Converter didn't create output file at {outputPath}");
                }

                Console.WriteLine($"File exists after conversion: {System.IO.File.Exists(outputPath)}");
                Console.WriteLine("Conversion succeeded.");

                model.ConversionResult = new ConversionResult
                {
                    Success = true,
                    InputFileName = inputFileName,
                    OutputFileName = outputFileName,
                    ConverterType = model.SelectedConverter,
                    OutputFilePath = "/downloads/" + outputFileName
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Conversion failed: {ex.Message}");
                model.ConversionResult = new ConversionResult
                {
                    Success = false,
                    ErrorMessage = $"Conversion failed: {ex.Message}"
                };
            }
            finally
            {
                // Clean up the temporary input file
                try
                {
                    if (System.IO.File.Exists(tempInputPath))
                    {
                        System.IO.File.Delete(tempInputPath);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error cleaning up temp file: {ex.Message}");
                }
            }

            return View("Index", model);
        }

        [HttpGet]
        public IActionResult Download(string fileName)
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "downloads", fileName);

            if (!System.IO.File.Exists(filePath))
            {
                return NotFound($"File not found: {fileName}");
            }

            // Determine content type based on file extension
            var contentType = "application/octet-stream"; // Default
            var extension = Path.GetExtension(fileName).ToLowerInvariant();

            switch (extension)
            {
                case ".pdf":
                    contentType = "application/pdf";
                    break;
                case ".docx":
                    contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                    break;
                case ".xlsx":
                    contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    break;
                case ".txt":
                    contentType = "text/plain";
                    break;
            }

            // Return file with the physical path
            return PhysicalFile(filePath, contentType, fileName);
        }
    }
}