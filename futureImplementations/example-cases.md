
### ✅ **DOCX to Other Formats:**  
1️. **DOCX to PDF:**  
   ```sh
   dotnet run docx2pdf ./Documents/Doc6.docx ./Documents/Doc6.pdf
   ```

2️. **DOCX to Text:**  
   ```sh
   dotnet run docx2txt ./Documents/test.docx ./Documents/test.txt
   ```
3️. **DOCX to Excel:**  
   ```sh
   dotnet run docx2excel ./Documents/test.docx ./Documents/test.xlsx
   ```

---

### ✅ **PDF to Other Formats:**  
4️. **PDF to DOCX:**  
   ```sh
   dotnet run pdf2docx ./Documents/sample.pdf ./Documents/sample.docx
   ```

5️. **PDF to Text:**  
   ```sh
   dotnet run pdf2txt ./Documents/sample.pdf ./Documents/sample.txt
   ```

---

### ✅ **HTML to Other Formats:**  
6️. **HTML to DOCX:**  
   ```sh
   dotnet run html2docx ./Documents/sample.html ./Documents/sample.docx
   ```
---

   **DOCX to HTML:**  
   ```sh
   dotnet run docx2html ./Documents/test.docx ./Documents/test.html
   ```

## **🛠 Future Improvements**

### **🚨 Edge Case Tests**  
❌ **Invalid Input Path:**  
   ```sh
   dotnet run docx2pdf ./Documents/missing.docx ./Documents/output.pdf
   ```
   _Expected: "Error: Input file './Documents/missing.docx' does not exist"_

❌ **Unsupported Converter Type errors to be fixed:**  
   ```sh
   dotnet run unknownConverter ./Documents/test.docx ./Documents/output.txt
   ```
   _Expected: "Error: Converter type 'unknownConverter' is not supported"_

   **HTML to PDF:**  
   ```sh
   dotnet run html2pdf ./Documents/sample.html ./Documents/sample.pdf
   ```
   _Expected: "Error: Converter type 'html2pdf' is not supported"_

   **Excel to DOCX:**  
   ```sh
   dotnet run excel2docx ./Documents/sample.xlsx ./Documents/sample.docx
   ```
   _Expected: "Error: Converter type 'excel2docx' is not supported"_

   **Excel to PDF:**  
   ```sh
   dotnet run excel2pdf ./Documents/sample.xlsx ./Documents/sample.pdf
   ```
   _Expected: "Error: Converter type 'excel2pdf' is not supported"_