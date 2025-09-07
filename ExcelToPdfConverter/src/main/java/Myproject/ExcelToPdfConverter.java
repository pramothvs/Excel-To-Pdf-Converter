package Myproject;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.*;

import java.io.*;
import java.util.Scanner;

public class ExcelToPdfConverter {
	 public static void main(String[] args) {
	        Scanner scanner = new Scanner(System.in);
	        String excelPath;
	        File excelFile;
	        // Get Excel file path
	        while (true) {
	            System.out.print("Enter Excel file path: ");
	            excelPath = scanner.nextLine();
	            excelFile = new File(excelPath);

	            if (!excelFile.exists() || !excelFile.isFile() || 
	               !(excelPath.endsWith(".xlsx") || excelPath.endsWith(".xls"))) {
	                System.out.println("Invalid file. Please enter a valid .xlsx file path.");
	            } else {
	                break; // valid, exit loop
	            }
	        }


	        // Get output PDF path
	        String pdfPath;
	        while (true) {
	            System.out.print("Enter output PDF path: ");
	            pdfPath = scanner.nextLine().trim();

	            // If user gives only folder, add default filename
	            File pdfFile = new File(pdfPath);
	            if (pdfFile.isDirectory()) {
	                pdfPath = pdfPath + File.separator + "output.pdf";
	                System.out.println("No file given, saving as: " + pdfPath);
	            }

	            if (pdfPath.toLowerCase().endsWith(".pdf")) {
	                try {
	                    // Try to create file 
	                    File testFile = new File(pdfPath);
	                    testFile.getParentFile().mkdirs();
	                    if (testFile.createNewFile() || testFile.exists()) {
	                        testFile.delete(); 
	                        break; 
	                    }
	                } catch (Exception e) {
	                    System.out.println("Cannot write to this location. Try again.");
	                }
	            } else {
	                System.out.println("Invalid file. Path must end with .pdf");
	            }
	        }
	        scanner.close();


	        try (InputStream inputStream = new FileInputStream(excelFile);
	             Workbook workbook = WorkbookFactory.create(inputStream)) {

	            Document document = new Document(PageSize.A4);
	            PdfWriter.getInstance(document, new FileOutputStream(pdfPath));
	            document.open();

	            for (int s = 0; s < workbook.getNumberOfSheets(); s++) {
	                Sheet sheet = workbook.getSheetAt(s);
	                if (sheet.getPhysicalNumberOfRows() == 0) continue;

	                int numCols = sheet.getRow(0).getLastCellNum();
	                PdfPTable pdfTable = new PdfPTable(numCols);
	                pdfTable.setWidthPercentage(100);
	                
	                // Set equal column widths
	                float[] columnWidths = new float[numCols];
	                for (int i = 0; i < numCols; i++) columnWidths[i] = 1f;
	                pdfTable.setWidths(columnWidths);

	                for (int r = 0; r <= sheet.getLastRowNum(); r++) {
	                    Row row = sheet.getRow(r);
	                    if (row == null) continue;

	                    for (int c = 0; c < row.getLastCellNum(); c++) {
	                        Cell cell = row.getCell(c);

	                        if (isCellCoveredByMergedRegion(sheet, r, c)) {
	                            continue; 
	                        }

	                        if (cell != null) {
	                            // Get Excel style + font
	                            CellStyle style = cell.getCellStyle();
	                            org.apache.poi.ss.usermodel.Font excelFont = workbook.getFontAt(style.getFontIndexAsInt());

	                            int fontStyle = Font.NORMAL;
	                            if (excelFont.getBold()) fontStyle |= Font.BOLD;
	                            if (excelFont.getItalic()) fontStyle |= Font.ITALIC;

	                            Font pdfFont = new Font(
	                                    Font.FontFamily.HELVETICA,
	                                    excelFont.getFontHeightInPoints() > 0 ? excelFont.getFontHeightInPoints() : 11,
	                                    fontStyle,
	                                    BaseColor.BLACK // temporarily keep text black
	                            );

	                            // Create phrase with text
	                            String text = getCellValue(cell);
	                            Phrase phrase = new Phrase(text, pdfFont);
	                            PdfPCell pdfCell = new PdfPCell(phrase);

	                            // Alignment
	                            switch (style.getAlignment()) {
	                                case CENTER: pdfCell.setHorizontalAlignment(Element.ALIGN_CENTER); break;
	                                case RIGHT:  pdfCell.setHorizontalAlignment(Element.ALIGN_RIGHT); break;
	                                default:     pdfCell.setHorizontalAlignment(Element.ALIGN_LEFT); break;
	                            }
	                            pdfCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
	                            
	                         // Wrap text and padding
	                            pdfCell.setNoWrap(false);
	                            pdfCell.setPadding(5f);
	                            pdfCell.setUseAscender(true);
	                            pdfCell.setUseDescender(true);

	                            // Handle merged regions
	                            CellRangeAddress mergedRegion = getMergedRegion(sheet, r, c);
	                            if (mergedRegion != null) {
	                                int colspan = mergedRegion.getLastColumn() - mergedRegion.getFirstColumn() + 1;
	                                int rowspan = mergedRegion.getLastRow() - mergedRegion.getFirstRow() + 1;
	                                pdfCell.setColspan(colspan);
	                                pdfCell.setRowspan(rowspan);
	                            }

	                            pdfTable.addCell(pdfCell);
	                        }
	                    }
	                }
	                document.add(pdfTable);
	                document.newPage();
	            }

	            document.close();
	            System.out.println("PDF created: " + pdfPath);

	        } catch (IOException | DocumentException e) {
	            System.err.println("Conversion failed: " + e.getMessage());
	        }
	    }

	 private static String getCellValue(Cell cell) {
		    switch (cell.getCellType()) {
		        case STRING:
		            return cell.getStringCellValue();
		        case NUMERIC:
		        	if (DateUtil.isCellDateFormatted(cell)) {
		                return cell.getDateCellValue().toString();
		            } 
		        	else {
		                double num = cell.getNumericCellValue();
		                if (num == (long)num) {
		                    return String.valueOf((long)num); // show as integer if no decimal
		                } 
		                else {
		                    return String.valueOf(num);}
		            }
		        case BOOLEAN:
		            return String.valueOf(cell.getBooleanCellValue());
		        case FORMULA:
		            return cell.getCellFormula();
			default:
				return "";
		    }
		}
	 private static CellRangeAddress getMergedRegion(Sheet sheet, int row, int column) {
		    for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
		        CellRangeAddress region = sheet.getMergedRegion(i);
		        if (region.isInRange(row, column)) {
		            return region;
		        }
		    }
		    return null;
		}

	 private static boolean isCellCoveredByMergedRegion(Sheet sheet, int row, int column) {
		    for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
		    	CellRangeAddress region = sheet.getMergedRegion(i);
		        if (region.isInRange(row, column)) {
		            if (region.getFirstRow() != row || region.getFirstColumn() != column) {
		                return true;}
		            }
		        }
		    return false;
		    }
}
