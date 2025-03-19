using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;

namespace DocumentConverterApplication.Controllers
{
    public class HomeController : Controller
    {
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
        public async Task<IActionResult> ConvertDocument(ConverterViewModel model)
        {
            if (!ModelState.IsValid || model.UploadedFile == null || string.IsNullOrEmpty(model.SelectedConverter))
            {
                model.IsModelStateValid = false;
                return View("Index", model);
            }

            // Add real async operation - for example if you're processing the file
            await Task.Run(() =>
            {
                // Simulate file processing work
                System.Threading.Thread.Sleep(100);
            });

            // Simulate conversion process
            var inputFileName = Path.GetFileName(model.UploadedFile.FileName);
            var outputFileName = "converted_" + inputFileName;
            var outputPath = "/downloads/" + outputFileName;

            model.ConversionResult = new ConversionResult
            {
                Success = true,
                InputFileName = inputFileName,
                OutputFileName = outputFileName,
                ConverterType = model.SelectedConverter,
                OutputFilePath = outputPath
            };

            model.IsModelStateValid = true;
            return View("Index", model);
        }
    }
}