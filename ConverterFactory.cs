public class ConverterFactory
{
    public static DocumentConverter CreateConverter(string converterType)
    {
        switch (converterType.ToLower())
        {
            // DOCX converters
            case "docx2pdf":
                return new DocxToPdfConverter();
            case "docx2html":
                return new DocxToHtmlConverter();
            case "docx2txt":
                return new DocxToTextConverter();
            case "docx2excel":
                return new DocxToExcelConverter();
                
            // PDF converters
            case "pdf2docx":
                return new PdfToDocxConverter();
            case "pdf2html":
                return new PdfToHtmlConverter();
            case "pdf2txt":
                return new PdfToTextConverter();
                
            // HTML converters
            case "html2docx":
                return new HtmlToDocxConverter();
            case "html2pdf":
                return new HtmlToPdfConverter();
                
            // Excel converters
            case "excel2docx":
                return new ExcelToDocxConverter();
            case "excel2pdf":
                return new ExcelToPdfConverter();
                
            default:
                throw new ArgumentException($"Unsupported converter type: {converterType}");
        }
    }
}
