package excelOperations;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelToHashMap {

	public static void main(String[] args) throws IOException {
		
		String excelFilePath = ".\\dataFiles\\student.xlsx";
		FileInputStream inputStream = new FileInputStream(excelFilePath);
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workbook.getSheet("Student Data");
			
		int rows = sheet.getLastRowNum();
		HashMap<String,String> data = new HashMap<String,String>();
		
		for(int r=0;r<=rows;r++) {
			String key = sheet.getRow(r).getCell(0).getStringCellValue();
			String value = sheet.getRow(r).getCell(1).getStringCellValue();
			data.put(key, value);
		}

		for(Map.Entry<String,String> entry:data.entrySet())
		{
			System.out.println(entry.getKey() + " " + entry.getValue());
		}
		
		System.out.println("Done.!!");

		
	}

}
