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

			//setupMainSheet(mainSheet);

			setUpEdittedData(edittedData);
			//int typeIndex = Arrays.asList(firstRowMainSheet).indexOf("C_Type");
			for (int i = 0; i < mainSheet.getPhysicalNumberOfRows(); i++) {
				if (isACoOp(i, 5	) || isAnInternship(i, 5) || isResearch(i, 5)
						|| isPartTimeWork(i, 5)) {
					System.out.println("works");


					// works
					// System.out.println("Row " + i + "is yes");
					// find what registration class they need the pair for
					// works
					//int regIndex = Arrays.asList(firstRowMainSheet).indexOf("RegCode");
					//String regNum = wholeSheet[i][regIndex];
					//System.out.println(i + " " + regIndex + " " + regNum);

				} else {
					System.out.println("not a coop");
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

	public static void setupMainSheet(XSSFSheet mainSheet) {
		/*
		int i = 0, j = 0;
		wholeSheet = new String[mainSheet.getPhysicalNumberOfRows()][47];

		for (Row row : mainSheet) {
			for (Cell cell : row) {
		          int c = row.getCell(cell, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

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
		*/
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
		// System.out.println(wholeSheet[row][col]);
		if(mainSheet.getRow(row).getCell(col).toString().equals("1.0")) {
			//System.out.println("row " + row + " col " + col + " YES IS A CO OP");
			return true;
		}		
		//System.out.println(mainSheet.getRow(row).getCell(col).toString());
		//System.out.println("row " + row + " col " + col + " NO IS NOT A CO OP");

		return false;
	}
	

	public static boolean isAnInternship(int row, int col) {
		// works
		// System.out.println(wholeSheet[row][col]);
		if(mainSheet.getRow(row).getCell(col).toString().equals("2.0")) {
			// System.out.println("row " + row + " col " + col + " YES IS AN INTERNSHIP");
			return true;
		}
		// System.out.println("row " + row + " col " + col + " NO IS NOT AN
		// INTERNSHIP");

		return false;
	}

	public static boolean isResearch(int row, int col) {
		// works
		// System.out.println(wholeSheet[row][col]);
		if(mainSheet.getRow(row).getCell(col).toString().equals("4.0")) {
			// System.out.println("row " + row + " col " + col + " YES IS RESEARCH");
			return true;
		}
		// System.out.println("row " + row + " col " + col + " NO IS NOT RESEARCH");

		return false;
	}

	public static boolean isPartTimeWork(int row, int col) {
		// works
		// System.out.println(wholeSheet[row][col]);
		if(mainSheet.getRow(row).getCell(col).toString().equals("7.0")) {
			// System.out.println("row " + row + " col " + col + " YES IS PART TIME WORK");
			return true;
		}
		// System.out.println("row " + row + " col " + col + " NO IS NOT PART TIME
		// WORK");

		return false;
	}
	
}
