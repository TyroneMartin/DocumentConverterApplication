using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.IO;
using System;
using DocumentConverterApplication.Models;
using DocumentConverterApplication.Services;
using Microsoft.Extensions.Logging;
using ConverterLibrary;

namespace DocumentConverterApplication.Controllers
{
    public class HomeController : Controller
    {
        private readonly ITempFileService _tempFileService;
        private readonly ILogger<HomeController> _logger;

        public HomeController(ITempFileService tempFileService, ILogger<HomeController> logger)
        {
            _tempFileService = tempFileService;
            _logger = logger;
        }

        public IActionResult Index()
        {
            var model = new ConverterViewModel
            {
                AvailableConverters = new Dictionary<string, List<string>>
                {
                    { "Text", new List<string> { "txt to pdf", "txt to docx" } },
                    { "Document", new List<string> { "docx to pdf", "pdf to docx", "docx to excel", "docx to html" } },
                    { "Spreadsheet", new List<string> { "excel to pdf", "csv to xlsx" } }
                }
            };
            return View(model);
        }

        [HttpPost]
        public IActionResult ConvertDocument(ConverterViewModel model)
        {
            _logger.LogInformation("Starting document conversion process");

            // Repopulate converters
            model.AvailableConverters = new Dictionary<string, List<string>>
            {
                { "Text", new List<string> { "txt to pdf", "txt to docx" } },
                { "Document", new List<string> { "docx to pdf", "pdf to docx", "docx to excel", "docx to html" } },
                { "Spreadsheet", new List<string> { "excel to pdf", "csv to xlsx" } }
            };

            // Detailed logging
            _logger.LogInformation($"Uploaded File: {model.UploadedFile?.FileName}");
            _logger.LogInformation($"Selected Converter: {model.SelectedConverter}");

            // Validate inputs
            if (!ModelState.IsValid || model.UploadedFile == null)
            {
                _logger.LogWarning("Invalid model state or no file uploaded");
                model.IsModelStateValid = false;
                return View("Index", model);
            }

            try
            {
                // Create directories with detailed logging
                var tempDir = _tempFileService.GetTempDirectory();
                var downloadsDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "downloads");
                
                _logger.LogInformation($"Temp Directory: {tempDir}");
                _logger.LogInformation($"Downloads Directory: {downloadsDir}");

                Directory.CreateDirectory(tempDir);
                Directory.CreateDirectory(downloadsDir);

                // Generate unique file paths
                var inputFileName = Path.GetFileName(model.UploadedFile.FileName);
                var tempInputPath = Path.Combine(tempDir, $"{Guid.NewGuid()}{Path.GetExtension(inputFileName)}");
                
                // Save uploaded file
                using (var stream = new FileStream(tempInputPath, FileMode.Create))
                {
                    model.UploadedFile.CopyTo(stream);
                }

                // Normalize converter type for factory
                var normalizedConverterType = model.SelectedConverter
                    .Replace(" ", "")
                    .ToLower()
                    .Replace("to", "2");

                // Prepare output path
                var outputFileName = $"converted_{Path.GetFileNameWithoutExtension(inputFileName)}_{Guid.NewGuid()}.{GetOutputExtension(normalizedConverterType)}";
                var outputPath = Path.Combine(downloadsDir, outputFileName);

                _logger.LogInformation($"Converter Type: {normalizedConverterType}");
                _logger.LogInformation($"Input Path: {tempInputPath}");
                _logger.LogInformation($"Output Path: {outputPath}");

                // Perform conversion
                var converter = ConverterFactory.CreateConverter(normalizedConverterType);
                converter.ConvertWithValidation(tempInputPath, outputPath);

                // Set conversion result
                model.ConversionResult = new ConversionResult
                {
                    Success = true,
                    InputFileName = inputFileName,
                    OutputFileName = outputFileName,
                    OutputFilePath = $"/downloads/{outputFileName}",
                    ConverterType = model.SelectedConverter
                };

                // Clean up temp file
                System.IO.File.Delete(tempInputPath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Conversion failed");
                model.ConversionResult = new ConversionResult
                {
                    Success = false,
                    ErrorMessage = $"Conversion failed: {ex.Message}"
                };
            }

            return View("Index", model);
        }

        // Helper method to determine output extension
        private string GetOutputExtension(string converterType)
        {
            return converterType switch
            {
                "docx2pdf" => "pdf",
                "docx2html" => "html",
                "docx2txt" => "txt",
                "docx2excel" => "xlsx",
                "pdf2docx" => "docx",
                "pdf2txt" => "txt",
                "html2docx" => "docx",
                _ => "converted"
            };
        }
    }
}