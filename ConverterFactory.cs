using DocumentFormat.OpenXml.Spreadsheet;

namespace ConverterLibrary
{
    public class ConverterFactory
    {
        public static DocumentConverter CreateConverter(string converterType) =>
            converterType.ToLower() switch
            {
                // DOCX converters
                "docx2pdf" => new DocxToPdfConverter(),
                "docx2html" => new DocxToHtmlConverter(),
                "docx2txt" => new DocxToTextConverter(),
                "docx2excel" => new DocxToExcelConverter(),

                // PDF converters
                "pdf2docx" => new PdfToDocxConverter(),
                // "pdf2html" => new PdfToHtmlConverter(),
                "pdf2txt" => new PdfToTextConverter(),

                // HTML converters
                "html2docx" => new HtmlToDocxConverter(),

                // Future improvements

                // "html2pdf" => new HtmlToPdfConverter(),

                // Excel converters
                // "excel2docx" => new ExcelToDocxConverter(),
                // "excel2pdf" => new ExcelToPdfConverter(),                

                _ => throw new ArgumentException($"Unsupported converter type: {converterType}")
            };
    }
}
