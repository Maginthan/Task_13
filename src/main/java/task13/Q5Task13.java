package task13;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Q5Task13 {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		// Opening Excel file and open the first sheet
		XSSFWorkbook book = new XSSFWorkbook(
				"/Users/maginthangr/eclipse-workspace/Task13/src/main/java/task13/ReadFile.xlsx");
		XSSFSheet sheet = book.getSheetAt(0);

		// Getting the row count and column count size
		int rowCount = sheet.getLastRowNum();
		int columnCount = sheet.getRow(0).getLastCellNum();

		String[][] data = new String[rowCount][columnCount];

		System.out.println("Reading values from the Excel file ReadFile.xlsx");
		System.out.println("--------------------------------------");

		// Get into row
		for (int i = 1; i <= rowCount; i++) {
			XSSFRow row = sheet.getRow(i);
			System.out.println();

			// Get into cell
			for (int j = 0; j < columnCount; j++) {
				XSSFCell cell = row.getCell(j);

				// Getting the cell value and storing it to the data variable
				data[i - 1][j] = cell.getStringCellValue();
				System.out.print(cell.getStringCellValue() + " ");
			}

		}
		// Closing the Excel file
		book.close();
	}

}
