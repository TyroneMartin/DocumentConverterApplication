using Microsoft.AspNetCore.Http;
using System.Collections.Generic;

namespace DocumentConverterApplication.Models
{
    public class ConverterViewModel
    {
        public Dictionary<string, List<string>> AvailableConverters { get; set; } = new Dictionary<string, List<string>>();
        public string? SelectedConverter { get; set; }
        public IFormFile? UploadedFile { get; set; }
        public ConversionResult? ConversionResult { get; set; }
        public bool IsModelStateValid { get; set; } = true;
    }
}