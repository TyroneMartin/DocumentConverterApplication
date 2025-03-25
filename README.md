Below is an updated README.md tailored for your ASP.NET Core document converter web‑app. It details the web‑focused workflow, the technologies used, setup instructions, troubleshooting tips, and useful reference links.

---

```markdown
# Document Converter Web Application

A web‑based application built with ASP.NET Core for converting documents between various formats—such as DOCX, PDF, HTML, and Excel. The application leverages powerful libraries like iTextSharp, Open XML SDK, NPOI, and HtmlAgilityPack on the server side, while also providing a modern web interface for file uploads and downloads.

## Overview

The Document Converter Web Application allows users to upload files via their browser, have them converted on the server, and then download the converted output. The application supports conversions including:

- **DOCX Converters:**
  - `docx2pdf`: Convert DOCX to PDF.
  - `docx2html`: Convert DOCX to HTML.
  - `docx2txt`: Convert DOCX to plain text.
  - `docx2excel`: Convert DOCX to Excel.

- **PDF Converters:**
  - `pdf2docx`: Convert PDF to DOCX.
  - `pdf2txt`: Convert PDF to plain text.

- **HTML Converters:**
  - `html2docx`: Convert HTML to DOCX.

Future improvements include additional conversion types such as HTML to PDF and Excel conversions.

## Demo

Watch the [Software Demo Video](https://youtu.be/WVFQVa-WqDo) to see the application in action.

## Setup

1. **Clone the Repository**  
   ```bash
   git clone https://github.com/TyroneMartin/DocumentConverterApplication
   cd DocumentConverterApplication
   ```

2. **Install .NET SDK**  
   Download and install the .NET SDK from the [official .NET Downloads](https://dotnet.microsoft.com/download) page. Verify the installation by running:  
   ```bash
   dotnet --version
   ```

3. **Install Required NuGet Packages**  
   Restore and install the necessary NuGet packages:
   ```bash
   dotnet add package iTextSharp
   dotnet add package DocumentFormat.OpenXml
   dotnet add package NPOI
   dotnet add package HtmlAgilityPack
   dotnet add package itext7.bouncy-castle-adapter
   ```

4. **Build the Project**  
   Build the project to ensure that all dependencies are correctly implemented:
   ```bash
   dotnet build
   ```

5. **Run the Application**  
   Start the web application using:
   ```bash
   dotnet run
   ```
   The app will be available at [http://localhost:5000](http://localhost:5000).

## Web Application Workflow

- **File Upload & Conversion:**  
  Users upload a file and select a conversion type via the web form. The server processes the file using the appropriate converter (invoked via the `ConverterFactory`), saves the converted output in the `wwwroot/downloads` folder on the host machine, and provides a download link.

- **File Storage:**  
  The converted files are stored on the server (host machine) under `wwwroot/downloads`. They remain on the server until downloaded by the user or removed via a cleanup process.

## Technologies & Tools

- **Backend Framework:**  
  ASP.NET Core with MVC architecture for a robust and scalable web application.

- **Programming Language:**  
  C#

- **Key Libraries:**  
  - **iTextSharp / iText7:** For PDF conversion and manipulation.  
  - **DocumentFormat.OpenXml:** For working with DOCX files and other Office formats.  
  - **NPOI:** For Excel file creation and manipulation.  
  - **HtmlAgilityPack:** For parsing and converting HTML content.

- **Client-Side Technologies:**  
  HTML, CSS, and JavaScript for the web interface. (Optional: JavaScript-based conversion libraries can be integrated in the future for client-side processing.)

- **Development Environment:**  
  - **.NET SDK:** For building and running the application.  
  - **Visual Studio Code:** Lightweight editor with C# support.  
  - **Git:** Version control for managing the codebase.  
  - **NuGet:** Package manager for .NET dependencies.

## Features

- **User-Friendly Interface:**  
  A clean and responsive web interface for file upload, conversion, and download.

- **Multi-Format Support:**  
  Convert between DOCX, PDF, HTML, and Excel formats.

- **Server-Side Conversion:**  
  Conversion operations are performed on the server using reliable libraries, with the output stored in the `wwwroot/downloads` folder.

- **Robust Error Handling:**  
  Detailed logging and error messages are provided via console output and user notifications.

- **Extensible Architecture:**  
  New converters can be added easily by extending the `DocumentConverter` base class.

## Troubleshooting

### Form Submission Issues

- **Ensure Model Binding:**  
  Verify that the file input and conversion selection use the correct `asp-for` attributes so that the model is populated when the form is submitted.

- **Check Folder Permissions:**  
  The `wwwroot/downloads` folder must exist on the host machine, and the application must have write permissions to this folder.

### HTTPS Redirection Warning

If you see a warning like:  
```
Failed to determine the https port for redirect.
```
This occurs when the HTTPS redirection middleware cannot determine the HTTPS port. You can:
- Configure HTTPS in your `launchSettings.json`, or
- Disable HTTPS redirection (remove or comment out `app.UseHttpsRedirection()` in `Startup.cs`) during local development.

### Build and Runtime Errors

- **Clean and Rebuild:**  
  If you encounter build errors, run:
  ```bash
  dotnet clean
  dotnet build
  ```
- **Check Console Logs:**  
  The application logs (via `Console.WriteLine` and ILogger) provide insights into the conversion process and any errors encountered.

## Useful Websites & References

- **.NET Documentation:** [https://docs.microsoft.com/en-us/dotnet/](https://docs.microsoft.com/en-us/dotnet/)
- **iTextSharp Documentation:** [https://itextpdf.com/](https://itextpdf.com/)
- **Open XML SDK Documentation:** [https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk](https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk)
- **NPOI GitHub Repository:** [https://github.com/nissl-lab/npoi](https://github.com/nissl-lab/npoi)
- **HtmlAgilityPack:** [https://html-agility-pack.net/](https://html-agility-pack.net/)

## Time Spent

- **Development:** 25-30 hours  
- **Testing and Debugging:** 5-10 hours  
- **Documentation:** 2-3 hours
