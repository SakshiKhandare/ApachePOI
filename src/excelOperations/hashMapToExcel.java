package excelOperations;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class hashMapToExcel {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Student Data");
		
		Map<String,String> data = new HashMap<String, String>();
		data.put("101", "Alice");
		data.put("102", "Bob");
		data.put("103", "Alex");
		data.put("104", "Jim");
		data.put("105", "Kevin");
		
		int rowNo = 0;
		
		for(Map.Entry<String,String>  entry:data.entrySet())
		{
			XSSFRow row = sheet.createRow(rowNo++);
			row.createCell(0).setCellValue((String)entry.getKey());
			row.createCell(1).setCellValue((String)entry.getValue());
		}
		
		FileOutputStream fos = new FileOutputStream(".\\dataFiles\\student.xlsx");
		workbook.write(fos);
		fos.close();
		System.out.println("Done.!!");

		
	}

}

































