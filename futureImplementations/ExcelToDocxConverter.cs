using System;
using System.IO;
using System.Text; 
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class ExcelToDocxConverter : DocumentConverter
{
    public override string ExpectedInputExtension => ".xlsx";
    
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (Excel) to {outputPath} (DOCX)...");
        EnsureDirectoryExists(outputPath);

        // Read Excel file
        IWorkbook workbook;
        using (FileStream fs = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        {
            workbook = new XSSFWorkbook(fs);
        }
        
        // Create DOCX file
        using (WordprocessingDocument docx = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
        {
            // Add a main document part
            MainDocumentPart mainPart = docx.AddMainDocumentPart();
            
            // Create document structure
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());
            
            // Process each sheet
            for (int sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                
                // Add sheet name as heading
                Paragraph headingPara = body.AppendChild(new Paragraph());
                Run headingRun = headingPara.AppendChild(new Run());
                RunProperties headingProps = headingRun.AppendChild(new RunProperties());
                headingProps.AppendChild(new Bold());
                headingProps.AppendChild(new FontSize() { Val = "28" });
                headingRun.AppendChild(new Text($"Sheet: {sheet.SheetName}"));
                
                // Process each row
                for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    if (row != null)
                    {
                        Paragraph rowPara = body.AppendChild(new Paragraph());
                        StringBuilder rowText = new StringBuilder();
                        
                        for (int cellIndex = 0; cellIndex < row.LastCellNum; cellIndex++)
                        {
                            ICell cell = row.GetCell(cellIndex);
                            if (cell != null)
                            {
                                string cellValue = cell.ToString();
                                rowText.Append(cellValue).Append("\t");
                            }
                        }
                        
                        Run rowRun = rowPara.AppendChild(new Run());
                        rowRun.AppendChild(new Text(rowText.ToString()));
                    }
                }
                
                // Add separator between sheets
                body.AppendChild(new Paragraph());
            }
            
            mainPart.Document.Save();
        }

        Console.WriteLine("Conversion complete!");
    }
}