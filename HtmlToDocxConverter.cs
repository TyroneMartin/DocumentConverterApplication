using System;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;

public class HtmlToDocxConverter : DocumentConverter
{
    public override string ExpectedInputExtension => ".html";
    
    public override void Convert(string inputPath, string outputPath)
    {
        Console.WriteLine($"Converting {inputPath} (HTML) to {outputPath} (DOCX)...");
        EnsureDirectoryExists(outputPath);

        // Load HTML document
        HtmlDocument htmlDoc = new HtmlDocument();
        htmlDoc.Load(inputPath);
        
        // Create DOCX file
        using (WordprocessingDocument docx = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
        {
            // Add a main document part
            MainDocumentPart mainPart = docx.AddMainDocumentPart();
            
            // Create document structure
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());
            
            // Process HTML body content
            HtmlNode bodyNode = htmlDoc.DocumentNode.SelectSingleNode("//body");
            if (bodyNode != null)
            {
                ProcessHtmlNode(bodyNode, body);
            }
            
            mainPart.Document.Save();
        }

        Console.WriteLine("Conversion complete!");
    }
    
    private void ProcessHtmlNode(HtmlNode node, Body body)
    {
        foreach (HtmlNode childNode in node.ChildNodes)
        {
            if (childNode.NodeType == HtmlNodeType.Text)
            {
                string text = childNode.InnerText.Trim();
                if (!string.IsNullOrWhiteSpace(text))
                {
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    run.AppendChild(new Text(text));
                }
            }
            else if (childNode.Name == "p")
            {
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text(childNode.InnerText.Trim()));
            }
            else if (childNode.Name == "h1" || childNode.Name == "h2" || childNode.Name == "h3")
            {
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                RunProperties runProps = run.AppendChild(new RunProperties());
                runProps.AppendChild(new Bold());
                runProps.AppendChild(new FontSize() { Val = "28" });
                run.AppendChild(new Text(childNode.InnerText.Trim()));
            }
            else if (childNode.HasChildNodes)
            {
                ProcessHtmlNode(childNode, body);
            }
        }
    }
}