# Excel to PDF Converter

## Project Description
This Java Maven project converts Excel files (.xlsx or .xls) into PDF format.  
It preserves merged cells, font styles, alignment, and wraps long text properly.  

---

## Requirements
1. Java JDK 8 or higher
2. Apache Maven 3.6 or higher
3. Dependencies (downloaded automatically via Maven):
   - Apache POI (for reading Excel files)
   - iTextPDF (for generating PDF files)
4. Supported Input: Excel files (.xls, .xlsx)
5. Output: PDF files
6. Example files provided in samples/input and samples/output


---

## Input & Output Format

*Input:*  
- Excel file path (.xlsx or .xls)  
- Example:  C:\Users\YourName\Documents\data.xlsx

*Output:*  
- PDF file path to save converted file (give file name with .pdf )
- Example:  C:\Users\YourName\Documents\output.pdf

 ##  If only a folder path is given, the program saves as output.pdf in that folder.

 ---
