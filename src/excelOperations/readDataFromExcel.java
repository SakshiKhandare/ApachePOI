package excelOperations;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class readDataFromExcel {

	public static void main(String[] args) throws IOException {
	
		String excelFilePath = ".\\dataFiles\\countries.xlsx";
		FileInputStream inputStream = new FileInputStream(excelFilePath);
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		// XSSFSheet sheet = workbook.getSheetAt(0);
		
		
		///////////////////       FOR LOOP METHOD       ///////////////////
		/*
		// returns last row num which equals num of rows
		int rows = sheet.getLastRowNum();
		
		// returns num of cells in one row which means num of columns
		int columns = sheet.getRow(1).getLastCellNum();
		
		// outer for loop for rows and inner for loop for columns
		for(int r=0;r<=rows;r++) {
			XSSFRow row = sheet.getRow(r);
			
			for(int c=0;c<columns;c++) {
				// returns particular cell
				XSSFCell cell = row.getCell(c);
				
				switch(cell.getCellType())
				{
				case STRING: System.out.print(cell.getStringCellValue()+" "); break;
				case NUMERIC: System.out.print(cell.getNumericCellValue()+" "); break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue()+" "); break;
				default: break;
				}
			}
			System.out.println();
		}
		*/
		
		///////////////////       ITERATOR METHOD       ///////////////////
		
		Iterator iterator = sheet.iterator();
		
		while(iterator.hasNext()) {
			XSSFRow row = (XSSFRow)iterator.next();
			Iterator cellIterator = row.cellIterator();
			
			while(cellIterator.hasNext()) {
				XSSFCell cell = (XSSFCell)cellIterator.next();
				switch(cell.getCellType())
				{
				case STRING: System.out.print(cell.getStringCellValue()); break;
				case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
				default: break;
				}
				System.out.println(" | ");
			}
			System.out.println();
		}
		
		workbook.close();
		inputStream.close();
	}
}



































