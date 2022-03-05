package excelOperations;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writeDataToExcel2 {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp Info");
		
		ArrayList<Object[]> empdata = new ArrayList<Object[]>();
		empdata.add(new Object[]{"EmpID","Name","Job"});
		empdata.add(new Object[]{101,"David","Engineer"});
		empdata.add(new Object[]{102,"Smith","Manager"});
		empdata.add(new Object[]{103,"Scott","HR"});
		
		int rowCount = 0;
		for(Object[] emp:empdata) {
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
