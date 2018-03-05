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
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import com.microsoft.schemas.office.visio.x2012.main.CellType;

public class SemesterRegistration {

	public static XSSFRow row1MainSheet, row1EdittedData;
	public static XSSFSheet mainSheet, edittedData;
	public static String[] firstRowMainSheet, firstRowEdittedData;

	public static void main(String[] args) {

		try {
			XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("IRAbridged.xlsx"));
			workbook.setMissingCellPolicy(MissingCellPolicy.CREATE_NULL_AS_BLANK);
			workbook.setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
			mainSheet = workbook.getSheetAt(0);
			edittedData = workbook.createSheet("edittedData");

			// setupMainSheet(mainSheet);

			setUpEdittedData(edittedData);
			// int typeIndex = Arrays.asList(firstRowMainSheet).indexOf("C_Type");
			for (int i = 0; i < mainSheet.getPhysicalNumberOfRows(); i++) {
				if (isACoOp(i, 5) || isAnInternship(i, 5) || isResearch(i, 5) || isPartTimeWork(i, 5)) {

				} else {
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

}
