package task13;

import java.io.FileOutputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Q1234Task13 {

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		// Creating a new Excel workbook
		XSSFWorkbook book = new XSSFWorkbook();

		// Creating a new sheet in the above created workbook
		XSSFSheet sheet = book.createSheet("Sheet1");

		// Data to be written in the workbook
		String[][] data = { { "Name", "Age", "Email" }, { "John Doe", "30", "john@test.com" },
				{ "Jane Doe", "28", "john@test.com" }, { "Bob Smith", "35", "jacky@example.com" },
				{ "Swapnil", "37", "swapnil@example.com" } };

		int rowCount = 0;

		// Outer loop to iterate over row
		for (String[] row1 : data) {
			XSSFRow row = sheet.createRow(rowCount++);

			int columnCount = 0;

			// Inner loop to iterate over cell
			for (String col : row1) {
				XSSFCell cell = row.createCell(columnCount++);
				cell.setCellValue(col);
			}

		}

		// Using try catch to handle FileNotFoundException
		try (FileOutputStream output = new FileOutputStream(
				"/Users/maginthangr/eclipse-workspace/Task13/src/main/java/task13/WriteFile.xlsx");) {
			book.write(output);

			// Closing the workbook
			book.close();

			System.out.println("Data were written successfully to the WriteFile.xlsx: ");
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
