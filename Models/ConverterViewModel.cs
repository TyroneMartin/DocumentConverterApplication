using Microsoft.AspNetCore.Http;
using System.Collections.Generic;

namespace DocumentConverterApplication.Models;
public class ConverterViewModel
{
    public IFormFile? UploadedFile { get; set; }
    public string? SelectedConverter { get; set; }
    public Dictionary<string, List<string>> AvailableConverters { get; set; } = new();
    public ConversionResult? ConversionResult { get; set; }
    public bool IsModelStateValid { get; set; } = true;
}