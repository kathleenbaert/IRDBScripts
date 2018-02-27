package scripts;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.schemas.office.visio.x2012.main.CellType;

public class SemesterRegistration {

	public static XSSFRow row1;
	public static XSSFSheet mainSheet, edittedData;
	public static String[] firstRow;
	public static String[][] wholeSheet;

	public static void main(String[] args) {

		try {
			XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("IRAbridged.xlsx"));
			mainSheet = workbook.getSheetAt(0);
			edittedData = workbook.createSheet("edittedData");

			performSetup(mainSheet);
			int typeIndex = Arrays.asList(firstRow).indexOf("C_Type");
			for (int i = 0; i < mainSheet.getPhysicalNumberOfRows(); i++) {
				if (isACoOp(i, typeIndex)) {
				}
			}

			// write out here
			FileOutputStream fileOut = new FileOutputStream("DataOut.xlsx");
			workbook.write(fileOut);
			fileOut.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public static void performSetup(XSSFSheet mainSheet) {
		// get first row:
		row1 = mainSheet.getRow(0);
		firstRow = new String[47];// num of columns across
		for (int i = 0; i < firstRow.length; i++) {
			Cell cell = row1.getCell(i);
			firstRow[i] = cell.getRichStringCellValue().getString();
		}
		// works
		// for(int i = 0; i < firstRow.length; i++){
		// System.out.println(firstRow[i]);
		// }

		// read in rest of the sheet:

		int i = 0, j = 0;
		wholeSheet = new String[mainSheet.getPhysicalNumberOfRows()][47];

		for (Row row : mainSheet) {
			for (Cell cell : row) {

				switch (cell.getCellTypeEnum().toString()) {
				case ("STRING"):
					wholeSheet[i][j] = cell.getRichStringCellValue().getString();
					j++;
					break;
				case ("NUMERIC"):
					wholeSheet[i][j] = Double.toString(cell.getNumericCellValue());
					j++;
					break;
				case ("BOOLEAN"):
					wholeSheet[i][j] = Boolean.toString(cell.getBooleanCellValue());
					j++;
					break;
				default:
					wholeSheet[i][j] = "ABORT MISSION!!!";// should never get here
					j++;
				}

			}
			i++;
			j = 0;
		}
		// works
		// System.out.println(Arrays.deepToString(wholeSheet));
	}

	public static boolean isACoOp(int row, int col) {
		//works
		//System.out.println(wholeSheet[row][col]);
		if (wholeSheet[row][col].equals("1.0")) {
			//System.out.println("row " + row + " col " + col + " YES IS A CO OP");
			return true;
		}
		//System.out.println("row " + row + " col " + col + " NO IS NOT A CO OP");

		return false;
	}
}
