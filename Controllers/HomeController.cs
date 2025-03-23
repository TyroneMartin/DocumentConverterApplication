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
            if (!ModelState.IsValid || model.UploadedFile == null || string.IsNullOrEmpty(model.SelectedConverter))
            {
                model.IsModelStateValid = false;
                return View("Index", model);
            }
            
            // Use the injected _tempFileService to get the temp directory.
            var tempDir = _tempFileService.GetTempDirectory();
            var inputFileName = Path.GetFileName(model.UploadedFile.FileName);
            var tempInputPath = Path.Combine(tempDir, Guid.NewGuid() + Path.GetExtension(inputFileName));
            
            using (var stream = new FileStream(tempInputPath, FileMode.Create))
            {
                model.UploadedFile.CopyTo(stream);
            }
            
            // Determine output file path
            var outputFileName = "converted_" + inputFileName;
            var outputPath = Path.Combine("wwwroot", "downloads", outputFileName);
            
            try
            {
                // Create and invoke the appropriate converter.
                var converter = ConverterLibrary.ConverterFactory.CreateConverter(
                    model.SelectedConverter.Replace(" ", "").ToLower()
                );
                converter.ConvertWithValidation(tempInputPath, outputPath);
                
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
                model.ConversionResult = new ConversionResult
                {
                    Success = false,
                    ErrorMessage = $"Conversion failed: {ex.Message}"
                };
            }
            
            return View("Index", model);
        }
    }
}
