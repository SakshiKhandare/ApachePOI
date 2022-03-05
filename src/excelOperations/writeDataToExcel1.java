package excelOperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writeDataToExcel1 {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp Info");
		
		Object empdata[][] = {
				{"EmpID","Name","Job"},
				{101,"David","Engineer"},
				{102,"Smith","Manager"},
				{103,"Scott","HR"},
				{104,"Bob","Analyst"},
				{105,"Alex","Lead"}
		};
		
		///////////////////       FOR LOOP METHOD       ///////////////////
		
		/* int rows = empdata.length;
		int cols = empdata[0].length;
		
		System.out.println(rows+" "+cols);
		
		for(int r=0;r<rows;r++) {
			XSSFRow row = sheet.createRow(r);
			
			for(int c=0;c<cols;c++) {
				XSSFCell cell = row.createCell(c);
				Object value = empdata[r][c];
				
				if(value instanceof String)
					cell.setCellValue((String)value);
				if(value instanceof Integer)
					cell.setCellValue((Integer)value);
				if(value instanceof Boolean)
					cell.setCellValue((Boolean)value);
			}
		}
		*/

		///////////////////       FOR EACH LOOP METHOD       ///////////////////
		
		int rowCount = 0;
		
		for(Object emp[]:empdata) {
			XSSFRow row  = sheet.createRow(rowCount++);
			int colCount = 0;
			for(Object value:emp) {
				XSSFCell cell = row.createCell(colCount++);
				
				if(value instanceof String)
					cell.setCellValue((String)value);
				if(value instanceof Integer)
					cell.setCellValue((Integer)value);
				if(value instanceof Boolean)
					cell.setCellValue((Boolean)value);
				
			}
		}
		
		String filePath = ".\\dataFiles\\employees.xlsx";
		FileOutputStream stream = new FileOutputStream(filePath);
		workbook.write(stream);
		stream.close();
		System.out.println("employees.xlsx file written successfully");
	
	
		
	
	}
	
}
