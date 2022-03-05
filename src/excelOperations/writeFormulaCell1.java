package excelOperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writeFormulaCell1 {

	public static void main(String[] args) throws IOException {
		
		//create new workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		// create new sheet named Numbers
		XSSFSheet sheet = workbook.createSheet("Numbers");
		// create 1 empty row inside Numbers sheet
		XSSFRow row = sheet.createRow(0);
		
		// create cells and set their value
		row.createCell(0).setCellValue(10);
		row.createCell(1).setCellValue(20);
		row.createCell(2).setCellValue(30);
		// creating formula cell
		row.createCell(3).setCellFormula("A1*B1*C1");
		
	
		// write data into file i.e. create new file at runtime
		FileOutputStream fos = new FileOutputStream(".\\dataFiles\\calc.xlsx"); 
		
		// creating this workbook into file system
		workbook.write(fos);
		
		fos.close();
		System.out.println("calc.xlsx created");
		
		
	}

}













































