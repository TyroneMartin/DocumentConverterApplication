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

            var tempDir = _tempFileService.GetTempDirectory();
            var inputFileName = Path.GetFileName(model.UploadedFile.FileName);
            var tempInputPath = Path.Combine(tempDir, Guid.NewGuid() + Path.GetExtension(inputFileName));

            Console.WriteLine($"Saving uploaded file to temporary path: {tempInputPath}");
            using (var stream = new FileStream(tempInputPath, FileMode.Create))
            {
                model.UploadedFile.CopyTo(stream);
            }

            var outputFileName = "converted_" + inputFileName;
            var outputPath = Path.Combine("wwwroot", "downloads", outputFileName);
            Console.WriteLine($"Output file will be: {outputPath}");

            try
            {
                var converter = ConverterLibrary.ConverterFactory.CreateConverter(
                    model.SelectedConverter.Replace(" ", "").ToLower()
                );
                Console.WriteLine($"Converter selected: {converter.GetType().Name}");
                converter.ConvertWithValidation(tempInputPath, outputPath);
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

            return View("Index", model);
        }

    }
}
