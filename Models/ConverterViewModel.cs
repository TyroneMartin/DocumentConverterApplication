public class ConverterViewModel
{
    public ConverterViewModel()
    {
        AvailableConverters = new Dictionary<string, List<string>>();
        SelectedConverter = string.Empty;
        UploadedFile = null!; // Using null-forgiving operator
        ConversionResult = new ConversionResult();
    }

    public Dictionary<string, List<string>> AvailableConverters { get; set; }
    public string SelectedConverter { get; set; }
    public IFormFile UploadedFile { get; set; }
    public ConversionResult ConversionResult { get; set; }
    
    // Property for validation check
    public bool IsModelStateValid { get; set; } = true;
}