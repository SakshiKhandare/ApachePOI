package excelOperations;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class readingPasswordProtectedExcel {

	public static void main(String[] args) throws IOException {

		FileInputStream fis = new FileInputStream(".\\dataFiles\\customers.xlsx");
		String password = "test123";

		// XSSFWorkbook workbook = new XSSFWorkbook(fis);
		// Workbook workbook = WorkbookFactory.create(fis,password);

		XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(fis, password);
		XSSFSheet sheet = workbook.getSheetAt(0);

		int rows = sheet.getLastRowNum();
		int columns = sheet.getRow(1).getLastCellNum();
		//System.out.println(rows + " " + columns);

		for (int r = 0; r <= rows; r++) {
			XSSFRow row = sheet.getRow(r);

			for (int c = 0; c < columns; c++) {
				XSSFCell cell = row.getCell(c);

				switch(cell.getCellType())
				{
				case STRING: System.out.print(cell.getStringCellValue()+" "); break;
				case NUMERIC: System.out.print(cell.getNumericCellValue()+" "); break;
				case BOOLEAN: System.out.print(cell.getBooleanCellValue()+" "); break;
				case FORMULA: System.out.println(cell.getNumericCellValue()+" "); break;
				default: break;
				}
			}
			System.out.println();
		}

		workbook.close();
		fis.close();
	}

}
