package scripts;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;

public class SemesterRegistration {

	public static XSSFRow row1MainSheet, row1EdittedData;
	public static XSSFSheet mainSheet, edittedData;
	public static String[] firstRowMainSheet, firstRowEdittedData;
	public static int edittedDataCurrRow;

	public static void main(String[] args) {

		try {
			XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("IRAbridged.xlsx"));
			mainSheet = workbook.getSheetAt(0);
			edittedData = workbook.createSheet("edittedData");
			workbook.setMissingCellPolicy(MissingCellPolicy.CREATE_NULL_AS_BLANK);
			workbook.setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

			setUpEdittedData(edittedData);
			edittedDataCurrRow = 1;
			mainLoop();

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

	public static void setUpEdittedData(XSSFSheet edittedData) {

		firstRowEdittedData = new String[] { "ID", "MUID", "TERM", "COMPANY_ID", "ACTIVITY", "SALARY", "CITY", "STATE",
				"COUNTRY", "REGID", "WORK_REG", "WORK_GRADE", "GRADING_REG", "GRADING_GRADE", "EMPLOYER_EVAL_DATE",
				"EMPLOYER_EVAL", "EMPLOYER_AUTH", "STUDENT_EVAL_DATE", "STUDENT_EVAL", "STUDENT_EVAL_DATE" };

		for (int i = 0; i < firstRowEdittedData.length; i++) {
			CellReference cr = new CellReference(0, i);
			int r = cr.getRow();
			int c = cr.getCol();
			row1EdittedData = edittedData.getRow(r);
			if (row1EdittedData == null)
				row1EdittedData = edittedData.createRow(r);
			Cell cell = row1EdittedData.getCell(c);
			cell.setCellValue(firstRowEdittedData[i]);
		}

	}

	public static void mainLoop() {
		for (int i = 0; i < mainSheet.getPhysicalNumberOfRows(); i++) {
			if (isACoOp(i, 5) || isAnInternship(i, 5) || isResearch(i, 5) || isPartTimeWork(i, 5)) {
				transferCoOpInfo(i);
				if (is3991(i, 14)) {
					int found = find3992(i, 14);
					if (verifyMUID(i, 1, found, 1)) {
						// transferReg(found);
					} else {
					}
				}
				if (is3993(i, 14)) {
					int found = find3994(i, 14);
					if (verifyMUID(i, 1, found, 1)) {
					} else {
					}

				}
				if (is4991(i, 14)) {
					int found = find4992(i, 14);
					if (verifyMUID(i, 1, found, 1)) {
					} else {
					}

				}
				if (is4993(i, 14)) {
					int found = find4994(i, 14);
					if (verifyMUID(i, 1, found, 1)) {
					} else {
					}

				}
			}
		}
	}

	public static boolean isACoOp(int row, int col) {
		// works
		if (mainSheet.getRow(row).getCell(col).toString().equals("1.0")) {
			// System.out.println("row " + row + " col " + col + " YES IS A CO OP");
			return true;
		}

		return false;
	}

	public static boolean isAnInternship(int row, int col) {
		// works
		if (mainSheet.getRow(row).getCell(col).toString().equals("2.0")) {
			// System.out.println("row " + row + " col " + col + " YES IS AN INTERNSHIP");
			return true;
		}

		return false;
	}

	public static boolean isResearch(int row, int col) {
		// works
		if (mainSheet.getRow(row).getCell(col).toString().equals("4.0")) {
			// System.out.println("row " + row + " col " + col + " YES IS RESEARCH");
			return true;
		}
		return false;
	}

	public static boolean isPartTimeWork(int row, int col) {
		// works
		if (mainSheet.getRow(row).getCell(col).toString().equals("7.0")) {
			// System.out.println("row " + row + " col " + col + " YES IS PART TIME WORK");
			return true;
		}

		return false;
	}

	public static boolean is3991(int row, int col) {
		if (mainSheet.getRow(row).getCell(col).toString().equals("3991.0")) {
			return true;
		}
		return false;

	}

	public static boolean is3993(int row, int col) {
		if (mainSheet.getRow(row).getCell(col).toString().equals("3993.0")) {
			return true;
		}
		return false;

	}

	public static boolean is4991(int row, int col) {
		if (mainSheet.getRow(row).getCell(col).toString().equals("4991.0")) {
			return true;
		}
		return false;

	}

	public static boolean is4993(int row, int col) {
		if (mainSheet.getRow(row).getCell(col).toString().equals("4993.0")) {
			return true;
		}
		return false;
	}

	public static int find3992(int row, int col) {
		int i = row;
		while (true) {
			// start search from current row, go down the line
			if (mainSheet.getRow(i).getCell(14).toString().equals("3992.0")) {
				// System.out.println("found it in: " + i);
				return i;
			}
			i++;
		}
	}

	public static int find3994(int row, int col) {
		int i = row;
		while (true) {
			// start search from current row, go down the line
			if (mainSheet.getRow(i).getCell(14).toString().equals("3994.0")) {
				// System.out.println("found it in: " + i);
				return i;
			}
			i++;
		}

	}

	public static int find4992(int row, int col) {
		int i = row;
		while (true) {
			// start search from current row, go down the line
			if (mainSheet.getRow(i).getCell(14).toString().equals("4992.0")) {
				// System.out.println("found it in: " + i);
				return i;
			}
			i++;
		}

	}

	public static int find4994(int row, int col) {
		int i = row;
		while (true) {
			// start search from current row, go down the line
			if (mainSheet.getRow(i).getCell(14).toString().equals("4994.0")) {
				// System.out.println("found it in: " + i);
				return i;
			}
			i++;
		}

	}

	public static boolean verifyMUID(int row1, int col1, int row2, int col2) {
		// System.out.println(mainSheet.getRow(row1).getCell(col1));
		// System.out.println(mainSheet.getRow(row2).getCell(col2));
		if (mainSheet.getRow(row1).getCell(col1).toString().equals((mainSheet.getRow(row2).getCell(col2).toString()))) {
			// System.out.println("yes accurate");
			return true;
		}
		return false;
	}

	public static void transferCoOpInfo(int srow) {
		// 0 id

		Row row = edittedData.getRow(edittedDataCurrRow);
		if (row == null) {
			row = edittedData.createRow(edittedDataCurrRow);
		}
		edittedData.getRow(edittedDataCurrRow).getCell(0)
				.setCellValue(mainSheet.getRow(srow).getCell(0).getNumericCellValue());
		// 1 MUID
		edittedData.getRow(edittedDataCurrRow).getCell(1)
				.setCellValue(mainSheet.getRow(srow).getCell(1).getNumericCellValue());
		// 2 term
		edittedData.getRow(edittedDataCurrRow).getCell(2)
				.setCellValue(mainSheet.getRow(srow).getCell(2).getNumericCellValue());
		// 3 company id
		edittedData.getRow(edittedDataCurrRow).getCell(3)
				.setCellValue(mainSheet.getRow(srow).getCell(3).getNumericCellValue());
		// 4 activity
		edittedData.getRow(edittedDataCurrRow).getCell(4)
				.setCellValue(mainSheet.getRow(srow).getCell(5).getNumericCellValue());
		// 5 salary
		edittedData.getRow(edittedDataCurrRow).getCell(5)
				.setCellValue(mainSheet.getRow(srow).getCell(6).getNumericCellValue());
		// 6 city
		edittedData.getRow(edittedDataCurrRow).getCell(6).setCellValue(mainSheet.getRow(srow).getCell(8).toString());
		// 7 state
		edittedData.getRow(edittedDataCurrRow).getCell(7).setCellValue(mainSheet.getRow(srow).getCell(9).toString());
		// 8 country
		edittedData.getRow(edittedDataCurrRow).getCell(8).setCellValue(mainSheet.getRow(srow).getCell(10).toString());
		// 9 regid
		edittedData.getRow(edittedDataCurrRow).getCell(9)
				.setCellValue(mainSheet.getRow(srow).getCell(11).getNumericCellValue());
		// 10 work_reg
		edittedData.getRow(edittedDataCurrRow).getCell(10)
				.setCellValue(mainSheet.getRow(srow).getCell(14).getNumericCellValue());

		// find work grade
		double temp = mainSheet.getRow(srow).getCell(14).getNumericCellValue();
		int c = 16;

		if (temp == 3991) {
			c = 16;
		}
		if (temp == 3993) {
			c = 17;
		}
		if (temp == 4991) {
			c = 18;
		}
		if (temp == 4993) {
			c = 19;
		}
		if (temp == 3992) {
			c = 20;
		}
		if (temp == 3994) {
			c = 21;
		}
		if (temp == 4992) {
			c = 22;
		}
		if (temp == 4994) {
			c = 23;
		}
		edittedData.getRow(edittedDataCurrRow).getCell(11).setCellValue(mainSheet.getRow(srow).getCell(c).toString());
		edittedDataCurrRow++;
	}

	public static void transferReg(int srow) {

		edittedData.getRow(edittedDataCurrRow).getCell(12)
				.setCellValue(mainSheet.getRow(srow).getCell(14).getNumericCellValue());
		edittedDataCurrRow++;
	}

	public static void transferGrade(int srow, int scol, int drow, int dcol) {

		edittedDataCurrRow++;
	}

}