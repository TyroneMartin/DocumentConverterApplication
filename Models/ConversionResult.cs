public class ConversionResult
{
    public ConversionResult()
    {
        InputFileName = string.Empty;
        OutputFileName = string.Empty;
        OutputFilePath = string.Empty;
        ConverterType = string.Empty;
        ErrorMessage = string.Empty;
    }

    public bool Success { get; set; }
    public string InputFileName { get; set; }
    public string OutputFileName { get; set; }
    public string OutputFilePath { get; set; }
    public string ConverterType { get; set; }
    public string ErrorMessage { get; set; }
}