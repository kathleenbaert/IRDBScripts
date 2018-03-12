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
			System.out.println("...set up editted data");
			edittedDataCurrRow = 1;
			transferLoop();
			System.out.println("...transfer loop");

			cleanUpLoop();

			System.out.println("...clean up loop");
			findDoubles();
			System.out.println("...find doubles");

			// write out here
			FileOutputStream fileOut = new FileOutputStream("DataOut.xlsx");
			workbook.write(fileOut);
			fileOut.close();
			System.out.println("_________________");
			System.out.println("PROGRAM COMPLETE");
			System.out.println("_________________");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public static void setUpEdittedData(XSSFSheet edittedData) {

		firstRowEdittedData = new String[] { "ID", "MUID", "TERM", "COMPANY_ID", "ACTIVITY", "SALARY", "CITY", "STATE",
				"COUNTRY", "REGID", "WORK_REG", "WORK_GRADE", "GRADING_REG", "GRADING_GRADE", "EMPLOYER_EVAL_DATE",
				"EMPLOYER_EVAL", "EMPLOYER_AUTH", "STUDENT_EVAL_DATE", "STUDENT_EVAL", "STUDENT_AUTH", "NOTES" };

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

	public static void transferLoop() {
		// for (int i = 1; i < mainSheet.getPhysicalNumberOfRows(); i++) {
		int i = 1;
		while (i < mainSheet.getPhysicalNumberOfRows()) {
			int last = endOfMUID(i);
			// System.out.println(i + " " + last);
			// System.out.println(mainSheet.getRow(i).getCell(1).toString() +" "+
			// mainSheet.getRow(i).getCell(1).toString());
			for (int j = i; j <= last; j++) {
				// System.out.println(j);
				if (is3991(j, 14)) {
					transferCoOpInfo(j);
					int found = find3992(j, 14, last);
					if(found == 0)
						break;
					transferWorkGrade(j, 16);
					transferGradingRegCredit(found, 20);
					transferEvals(j);
				}
				if (is3993(j, 14)) {
					transferCoOpInfo(j);
					int found = find3994(j, 14, last);
					if(found == 0)
						break;
					transferWorkGrade(j, 17);
					transferGradingRegCredit(found, 21);
					transferEvals(j);
				}
				if (is4991(j, 14)) {
					transferCoOpInfo(j);
					int found = find4992(j, 14, last);
					if(found == 0)
						break;
					transferWorkGrade(j, 18);
					transferGradingRegCredit(found, 22);
					transferEvals(j);
				}
				if (is4993(j, 14)) {
					transferCoOpInfo(j);
					int found = find4994(j, 14, last);
					if(found == 0)
						break;
					transferWorkGrade(j, 19);
					transferGradingRegCredit(found, 23);
					transferEvals(j);
				}

			}

			i = last + 1;
		}

	}

	public static int endOfMUID(int first) {
		int last = first + 1;
		if (!verifyMUID(first, 1, last, 1)) {// edge case, if only one entry for MUID
			return first;
		}
		while (verifyMUID(first, 1, last, 1)) {
			last++;
		}
		return last - 1;
	}

	public static void cleanUpLoop() {

		for (int i = 1; i < edittedData.getPhysicalNumberOfRows(); i++) {
			Replacements r = new Replacements();
			r.IRKeytoCheckmarqKey(i, edittedData); // works
			// for employer
			r.convertEvalAuthInits(i, 16, edittedData);
			// for student
			r.convertEvalAuthInits(i, 19, edittedData);
			// for employer
			r.convertEvalNoteItemID(i, 15, edittedData);
			// for students
			r.convertEvalNoteItemID(i, 18, edittedData);

			r.convertIRStudentActivityPlans(i, edittedData);

		}

	}

	public static void findDoubles() {
		int original = edittedData.getPhysicalNumberOfRows();
		for (int i = 0; i < original - 1; i++) {
			// 8 = registration ID
			if (edittedData.getRow(i).getCell(9).toString().equals(edittedData.getRow(i + 1).getCell(9).toString())) {
				//System.out.println(i + " duplicates!");
				Row r = edittedData.getRow(i + 1);
				edittedData.removeRow(r);
				i++;
			}
		}
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

	public static int find3992(int row, int col, int last) {
		// while (true) {
		while (row <= last) {
			// start search from current row, go down the line
			if (mainSheet.getRow(row).getCell(14).toString().equals("3992.0")) {
				// System.out.println("found it in: " + i);
				return row;
			}
			row++;
		}
		return 0;
	}

	public static int find3994(int row, int col, int last) {
		int i = row;
		while (row <= last) {
			// start search from current row, go down the line
			if (mainSheet.getRow(row).getCell(14).toString().equals("3994.0")) {
				// System.out.println("found it in: " + i);
				return row;
			}
			row++;
		}
		return 0;

	}

	public static int find4992(int row, int col, int last) {
		int i = row;
		while (row <= last) {
			// start search from current row, go down the line
			if (mainSheet.getRow(row).getCell(14).toString().equals("4992.0")) {
				// System.out.println("found it in: " + i);
				return row;
			}
			row++;
		}
		return 0;

	}

	public static int find4994(int row, int col, int last) {
		int i = row;
		while (row <= last) {
			// start search from current row, go down the line
			if (mainSheet.getRow(row).getCell(14).toString().equals("4994.0")) {
				// System.out.println("found it in: " + i);
				return row;
			}
			row++;
		}
		return 0;

	}

	public static boolean verifyMUID(int row1, int col1, int row2, int col2) {
		if (mainSheet.getRow(row2) == null) {// means end of sheet
			return false;
		}
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
		if (mainSheet.getRow(srow).getCell(1).getCellType() == Cell.CELL_TYPE_NUMERIC) {
			edittedData.getRow(edittedDataCurrRow).getCell(1)
					.setCellValue(mainSheet.getRow(srow).getCell(1).getNumericCellValue());
		}
		if (mainSheet.getRow(srow).getCell(1).getCellType() == Cell.CELL_TYPE_STRING) {
			edittedData.getRow(edittedDataCurrRow).getCell(1)
					.setCellValue(mainSheet.getRow(srow).getCell(1).getStringCellValue());
		}
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

	}

	public static void transferWorkGrade(int srow, int scol) {
		// srow is current row in main sheet, scol is dependent on what credit you need
		// to find
		Row row = edittedData.getRow(edittedDataCurrRow);
		if (row == null) {
			row = edittedData.createRow(edittedDataCurrRow);
		}
		// WORK_GRADE
		edittedData.getRow(edittedDataCurrRow).getCell(11)
				.setCellValue(mainSheet.getRow(srow).getCell(scol).getStringCellValue());
	}

	public static void transferGradingRegCredit(int srow, int scol) {
		// srow is current row in main sheet, scol is dependent on what credit you need
		Row row = edittedData.getRow(edittedDataCurrRow);
		if (row == null) {
			row = edittedData.createRow(edittedDataCurrRow);
		}
		// GRADING_REG
		edittedData.getRow(edittedDataCurrRow).getCell(12)
				.setCellValue(mainSheet.getRow(srow).getCell(14).getNumericCellValue());
		// GRADING_GRADE
		edittedData.getRow(edittedDataCurrRow).getCell(13)
				.setCellValue(mainSheet.getRow(srow).getCell(scol).getStringCellValue());

	}

	public static void transferEvals(int srow) {

		// EMPLOYER_EVAL_DATE
		edittedData.getRow(edittedDataCurrRow).getCell(14)
				.setCellValue(mainSheet.getRow(srow).getCell(26).getDateCellValue());
		// EMPLOYER_EVAL
		edittedData.getRow(edittedDataCurrRow).getCell(15)
				.setCellValue(mainSheet.getRow(srow).getCell(38).getNumericCellValue());
		// EMPLOYER_AUTH
		edittedData.getRow(edittedDataCurrRow).getCell(16)
				.setCellValue(mainSheet.getRow(srow).getCell(27).getStringCellValue());
		// STUDENT_EVAL_DATE
		edittedData.getRow(edittedDataCurrRow).getCell(17)
				.setCellValue(mainSheet.getRow(srow).getCell(24).getDateCellValue());
		// STUDENT_EVAL
		edittedData.getRow(edittedDataCurrRow).getCell(18)
				.setCellValue(mainSheet.getRow(srow).getCell(38).getNumericCellValue());
		// STUDENT_AUTH
		edittedData.getRow(edittedDataCurrRow).getCell(19)
				.setCellValue(mainSheet.getRow(srow).getCell(25).getStringCellValue());
		// notes
		edittedData.getRow(edittedDataCurrRow).getCell(20)
				.setCellValue(mainSheet.getRow(srow).getCell(28).getStringCellValue());
		edittedDataCurrRow++;
	}

}