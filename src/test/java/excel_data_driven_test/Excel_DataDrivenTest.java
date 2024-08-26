package excel_data_driven_test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.collections4.bag.SynchronizedSortedBag;
import org.apache.logging.log4j.*;

public class Excel_DataDrivenTest {

	private static Logger demologger = LogManager.getLogger(Excel_DataDrivenTest.class.getName());

	public static void main(String[] args) throws FileNotFoundException {

	}

	public ArrayList getData(String testcaseName) throws FileNotFoundException {

		ArrayList<String> a = new ArrayList<String>();

		// To make XSSFWorkbook knows where is excel data, create FileInputStream
		FileInputStream fis = new FileInputStream("C:\\Users\\sokoeurn chhay\\OneDrive\\Desktop\\testdatasample.xlsx");

		// To read data in excel, Create XSSFWorkbook
		XSSFWorkbook workBook = new XSSFWorkbook();

		// want to know how many worksheet in workbook
		int sheets = workBook.getNumberOfSheets();

		for (int i = 0; i < sheets; i++) {
			if (workBook.getSheetName(i).equalsIgnoreCase("testdatasample")) {
				XSSFSheet sheet = workBook.getSheetAt(i);

				// identify testcase column by scanning the entire 1st row. Sheet is collection
				// of rows
				Iterator<Row> rows = sheet.iterator();
				Row firstRow = rows.next();
				Iterator<Cell> cell = firstRow.cellIterator(); // Row is collection of cells

				int k = 0;
				int column = 0;

				while (cell.hasNext()) {
					Cell cellValue = cell.next();
					if (cellValue.getStringCellValue().equalsIgnoreCase("TESTCASE")) {

						column = k;
					}

					k++;
				}

				System.out.println(column);

				// onece column is identified then scan the entired test column to identify test
				// case row.
				while (rows.hasNext()) {
					Row r = rows.next();
					if (r.getCell(column).getStringCellValue().equalsIgnoreCase(testcaseName)) {
						// After you grap testcase row, pull all datat of that row feed into test.
						Iterator<Cell> cv = r.cellIterator();
						while (cv.hasNext()) {
							
							Cell c = cv.next();
							
							if(c.getCellType()==CellType.STRING) {
								a.add(c.getStringCellValue());
							} else {
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
							
							
						}
					}

				}
			}
		}
		
		return a;

	}
}
