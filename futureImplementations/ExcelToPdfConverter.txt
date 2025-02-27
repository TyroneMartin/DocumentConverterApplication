using System;
using System.IO;
using System.Text; 
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Kernel.Font; 

public class ExcelToPdfConverter : DocumentConverter
{
    public override string ExpectedInputExtension => ".xlsx";
    
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (Excel) to {outputPath} (PDF)...");
        EnsureDirectoryExists(outputPath);

        // Read Excel file
        IWorkbook workbook;
        using (FileStream fs = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        {
            workbook = new XSSFWorkbook(fs);
        }
        
        // Create PDF file
        using (PdfWriter writer = new PdfWriter(outputPath))
        {
            using (PdfDocument pdfDoc = new PdfDocument(writer))
            {
                using (iText.Layout.Document document = new iText.Layout.Document(pdfDoc))
                {
                    // Process each sheet
                    for (int sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
                    {
                        ISheet sheet = workbook.GetSheetAt(sheetIndex);
                        
                        // Add sheet name as heading
                        document.Add(new Paragraph($"Sheet: {sheet.SheetName}")
                            .SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD))
                            .SetFontSize(16));
                        
                        // Create table for sheet data
                        int maxColumns = GetMaxColumns(sheet);
                        Table table = new Table(maxColumns);
                        
                        // Process each row
                        for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                        {
                            IRow row = sheet.GetRow(rowIndex);
                            if (row != null)
                            {
                                for (int cellIndex = 0; cellIndex < maxColumns; cellIndex++)
                                {
                                    ICell cell = row.GetCell(cellIndex);
                                    string cellValue = cell != null ? cell.ToString() : "";
                                    table.AddCell(new Cell().Add(new Paragraph(cellValue)));
                                }
                            }
                        }
                        
                        document.Add(table);
                        
                        // Add space between sheets
                        if (sheetIndex < workbook.NumberOfSheets - 1)
                        {
                            document.Add(new Paragraph("\n"));
                        }
                    }
                }
            }
        }

        Console.WriteLine("Conversion complete!");
    }
    
    private int GetMaxColumns(ISheet sheet)
    {
        int maxColumns = 0;
        for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            IRow row = sheet.GetRow(rowIndex);
            if (row != null && row.LastCellNum > maxColumns)
            {
                maxColumns = row.LastCellNum;
            }
        }
        return Math.Max(1, maxColumns); 
    }
}