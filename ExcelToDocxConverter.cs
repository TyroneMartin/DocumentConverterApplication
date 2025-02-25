using System;
using System.IO;
using System.Text; 
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class ExcelToDocxConverter : DocumentConverter
{
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (Excel) to {outputPath} (DOCX)...");
        EnsureDirectoryExists(outputPath);

        // Open Excel file
        using (var fs = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
        {
            IWorkbook workbook = new XSSFWorkbook(fs);
            ISheet sheet = workbook.GetSheetAt(0);

            // Create Word document
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputPath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                // Add a main document part
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                
                // Create the document structure
                mainPart.Document = new Document();
                Body body = new Body();
                mainPart.Document.Append(body);

                // Add content from Excel to Word
                for (int i = 0; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        StringBuilder rowText = new StringBuilder();
                        for (int j = 0; j < row.LastCellNum; j++)
                        {
                            ICell cell = row.GetCell(j);
                            if (cell != null)
                            {
                                rowText.Append(cell.ToString() + "\t");
                            }
                        }

                        if (rowText.Length > 0)
                        {
                            Paragraph para = new Paragraph();
                            Run run = new Run();
                            Text text = new Text(rowText.ToString());
                            
                            run.Append(text);
                            para.Append(run);
                            body.Append(para);
                        }
                    }
                }
            }
        }

        Console.WriteLine("Conversion complete!");
    }
}