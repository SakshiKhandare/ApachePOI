package excelOperations;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writeFormulaCell2 {

	public static void main(String[] args) throws IOException {

		// opening file in read mode
		FileInputStream file = new FileInputStream(".\\dataFiles\\books.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
		sheet.getRow(7).getCell(2).setCellFormula("SUM(C2:C6)");

		file.close();

		FileOutputStream fos = new FileOutputStream(".\\dataFiles\\books.xlsx");
		workbook.write(fos);
		workbook.close();
		fos.close();
		System.out.println("Done");
	}

}
