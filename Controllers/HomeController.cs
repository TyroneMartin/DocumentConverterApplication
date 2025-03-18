using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.IO;
using Microsoft.AspNetCore.Http;
using System.Threading.Tasks;
using ConverterLibrary;
using System;
using DocumentConverterWebApp.Services;

namespace DocumentConverterApplication.Controllers
{
    public class HomeController : Controller
    {
        private readonly ITempFileService _tempFileService;

        public HomeController(ITempFileService tempFileService)
        {
            _tempFileService = tempFileService;
            // Initialize these properties to avoid the non-nullable warnings
            InputFileName = string.Empty;
            OutputFileName = string.Empty;
        }

        // These properties are causing the non-nullable warnings
        public string InputFileName { get; set; }
        public string OutputFileName { get; set; }

        public IActionResult Index()
        {
            var model = new ConverterViewModel
            {
                AvailableConverters = GetAvailableConverters()
            };
            return View(model);
        }

        [HttpPost]
        public async Task<IActionResult> ConvertDocument(ConverterViewModel model)
        {
            if (!ModelState.IsValid || model.UploadedFile == null)
            {
                model.AvailableConverters = GetAvailableConverters();
                return View("Index", model);
            }

            try
            {
                // Save the uploaded file to a temporary location
                string inputFilePath = await SaveUploadedFile(model.UploadedFile);
                string inputFileName = model.UploadedFile.FileName;

                // Determine the output file extension based on the converter
                string outputExtension = GetOutputExtension(model.SelectedConverter);
                string outputFileName = Path.GetFileNameWithoutExtension(inputFileName) + outputExtension;
                string outputFilePath = Path.Combine(_tempFileService.GetTempDirectory(), outputFileName);

                // Create and use the converter
                var converter = ConverterFactory.CreateConverter(model.SelectedConverter);
                converter.Convert(inputFilePath, outputFilePath);

                // Return the result
                model.ConversionResult = new ConversionResult
                {
                    Success = true,
                    InputFileName = inputFileName,
                    OutputFileName = outputFileName,
                    OutputFilePath = outputFilePath,
                    ConverterType = model.SelectedConverter
                };

                model.AvailableConverters = GetAvailableConverters();
                return View("Index", model);
            }
            catch (Exception ex)
            {
                ModelState.AddModelError("", $"Conversion failed: {ex.Message}");
                model.AvailableConverters = GetAvailableConverters();
                return View("Index", model);
            }
        }

        public IActionResult Download(string filePath, string fileName)
        {
            if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
            {
                return NotFound();
            }

            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            return File(fileBytes, GetContentType(filePath), fileName);
        }

        private async Task<string> SaveUploadedFile(IFormFile file)
        {
            var tempFilePath = Path.Combine(_tempFileService.GetTempDirectory(), file.FileName);
            _tempFileService.EnsureDirectoryExists(tempFilePath);

            using (var stream = new FileStream(tempFilePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            return tempFilePath;
        }

        private Dictionary<string, List<string>> GetAvailableConverters()
        {
            return new Dictionary<string, List<string>>
            {
                { "Word Document (DOCX)", new List<string> { "docx2pdf", "docx2html", "docx2txt", "docx2excel" } },
                { "PDF Document", new List<string> { "pdf2docx", "pdf2txt" } },
                { "HTML Document", new List<string> { "html2docx" } }
            };
        }

        private string GetOutputExtension(string converterType)
        {
            return converterType.ToLower() switch
            {
                "docx2pdf" => ".pdf",
                "docx2html" => ".html",
                "docx2txt" => ".txt",
                "docx2excel" => ".xlsx",
                "pdf2docx" => ".docx",
                "pdf2txt" => ".txt",
                "html2docx" => ".docx",
                _ => ".output"
            };
        }

        private string GetContentType(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            return extension switch
            {
                ".pdf" => "application/pdf",
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".html" => "text/html",
                ".txt" => "text/plain",
                ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                _ => "application/octet-stream"
            };
        }
    }

    public class ConverterViewModel
    {
        public Dictionary<string, List<string>> AvailableConverters { get; set; }
        public string SelectedConverter { get; set; }
        public IFormFile UploadedFile { get; set; }
        public ConversionResult ConversionResult { get; set; }
    }

    public class ConversionResult
    {
        public bool Success { get; set; }
        public string InputFileName { get; set; }
        public string OutputFileName { get; set; }
        public string OutputFilePath { get; set; }
        public string ConverterType { get; set; }
    }
}
