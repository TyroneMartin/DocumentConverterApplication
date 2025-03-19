namespace DocumentConverterApplication.Models
{
    public class ConversionResult
    {
        public bool Success { get; set; }
        public string? InputFileName { get; set; }
        public string? OutputFileName { get; set; }
        public string? OutputFilePath { get; set; }
        public string? ConverterType { get; set; }
        public string? ErrorMessage { get; set; }
    }
}