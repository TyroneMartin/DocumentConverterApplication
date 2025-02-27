
# Overview
### Document Converter Application

A C# application for converting documents between various formats, such as DOCX, PDF, HTML, and Excel. Built using .NET and libraries like iTextSharp, Open XML SDK, and NPOI.

## Demo
[Software Demo Video](https://www.youtube.com/watch?v=your-video-id)

## Setup

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/TyroneMartin/DocumentConverterApplication
   cd document-converter-application
   ```

2. **Install .NET SDK**:
   - Download and install the .NET SDK from [.NET Downloads](https://dotnet.microsoft.com/download).
   - Verify installation by running:
     ```bash
     dotnet --version
     ```

3. **Install Required NuGet Packages**:
   - Install the necessary NuGet packages for the project:
     ```bash
     dotnet add package iTextSharp
     dotnet add package DocumentFormat.OpenXml
     dotnet add package NPOI
     dotnet add package HtmlAgilityPack
     dotnet add package itext7.bouncy-castle-adapter
     ```

4. **Build the Project**:
   - Build the project to ensure all dependencies are implemented correctly 🙂.

     ```bash
     dotnet build
     ```

5. **Run the Application**:
   - Run the application with the appropriate arguments for conversion:
     ```bash
     dotnet run <converter> <inputPath> <outputPath>
     ```
   - Example:
     ```bash
     dotnet run docx2pdf document.docx output.pdf
     ```

## Available Converters

The application currently supports the following conversions:

- **DOCX Converters**:
  - `docx2pdf`: Convert DOCX to PDF.
  - `docx2html`: Convert DOCX to HTML.
  - `docx2txt`: Convert DOCX to plain text.
  - `docx2excel`: Convert DOCX to Excel.

- **PDF Converters**:
  - `pdf2docx`: Convert PDF to DOCX.
  - `pdf2txt`: Convert PDF to plain text.

- **HTML Converters**:
  - `html2docx`: Convert HTML to DOCX.

## Future Improvements

The following converters are planned for future implementation:

- **HTML to PDF**
- **Excel to DOCX**
- **Excel to PDF**

## Troubleshooting

### Missing NuGet Packages
If you encounter errors related to missing packages, ensure that all required NuGet packages are installed:
```bash
dotnet add package iTextSharp
dotnet add package DocumentFormat.OpenXml
dotnet add package NPOI
dotnet add package HtmlAgilityPack
dotnet add package itext7.bouncy-castle-adapter
```

### Build Errors
If the build fails, clean and rebuild the project:
```bash
dotnet clean
dotnet build
```

### File Path Issues
Ensure that the input and output file paths are correct and accessible. Use absolute paths if necessary.

## Development Environment

- **.NET SDK**: For building and running the application.
- **Visual Studio Code**: A lightweight code editor with C# support.
- **Git**: Version control system for managing the codebase.
- **NuGet**: Package manager for .NET.

## Tech Stack

- **C#**: Primary programming language.
- **.NET**: Framework for building the application.
- **iTextSharp**: Library for PDF manipulation.
- **DocumentFormat.OpenXml**: Microsoft's library for working with Office documents.
- **NPOI**: Library for Excel manipulation.
- **HtmlAgilityPack**: Library for HTML parsing and manipulation.

## Features

- **File input and output**: Improve the user experience by specifying input and output file paths.
- **Command-line interface**: Create a easy-to-use interface for running conversions.
- **Multi-Format Support**: Convert between DOCX, PDF, HTML, and Excel formats.
- **Flexible Conversion**: Supports both single and batch conversions.
- **Error Handling**: Robust error handling for invalid inputs or unsupported formats,etc.
- **Extensible Architecture**: Easily add new converters by extending the `DocumentConverter` base class.

## Requirements

- **.NET SDK 6.0 or higher**: Required for building and running the application.
- **NuGet**: For managing dependencies.
- **Supported Formats**: `Please note:` Input files must be in a supported format (DOCX, PDF, HTML, Excel).

## Useful Websites

- **.NET Documentation**: [https://docs.microsoft.com/en-us/dotnet/](https://docs.microsoft.com/en-us/dotnet/)
- **iTextSharp Documentation**: [https://itextpdf.com/](https://itextpdf.com/)
- **Open XML SDK Documentation**: [https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk](https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk)
- **NPOI Documentation**: [https://github.com/nissl-lab/npoi](https://github.com/nissl-lab/npoi)
- **HtmlAgilityPack Documentation**: [https://html-agility-pack.net/](https://html-agility-pack.net/)

## Time Spent

- **Development**: 25-30 hours
- **Testing and Debugging**: 5-10 hours
- **Documentation**: 2-3 hours