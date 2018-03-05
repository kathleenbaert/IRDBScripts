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
					if (is3991(i, 14)) {
						int found = find3992(i, 14);
						if(verifyMUID(i, 1, found, 1)) {
							System.out.println("MUID correct");
						}else {
							System.out.println("MUID NOT correct");
						}
					}
					if (is3993(i, 14)) {
						int found = find3994(i, 14);
						if(verifyMUID(i, 1, found, 1)) {
							System.out.println("MUID correct");
						}else {
							System.out.println("MUID NOT correct");
						}

					}
					if (is4991(i, 14)) {
						int found = find4992(i, 14);
						if(verifyMUID(i, 1, found, 1)) {
							System.out.println("MUID correct");
						}else {
							System.out.println("MUID NOT correct");
						}

					}
					if (is4993(i, 14)) {
						int found = find4994(i, 14);
						if(verifyMUID(i, 1, found, 1)) {
							System.out.println("MUID correct");
						}else {
							System.out.println("MUID NOT correct");
						}

					}
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
				//System.out.println("found it in: " + i);
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
					//System.out.println("found it in: " + i);
					return i;
				}
				i++;
			}
	 
	 } public static int find4992(int row, int col) {
			int i = row;
			while (true) {
				// start search from current row, go down the line
				if (mainSheet.getRow(i).getCell(14).toString().equals("4992.0")) {
					//System.out.println("found it in: " + i);
					return i;
				}
				i++;
			}
	  
	 } public static int find4994(int row, int col) {
			int i = row;
			while (true) {
				// start search from current row, go down the line
				if (mainSheet.getRow(i).getCell(14).toString().equals("4994.0")) {
					//System.out.println("found it in: " + i);
					return i;
				}
				i++;
			}
	  
	 } 
	  
	 public static boolean verifyMUID(int row1, int col1, int row2, int col2) {
		 //System.out.println(mainSheet.getRow(row1).getCell(col1));
		 //System.out.println(mainSheet.getRow(row2).getCell(col2));
		 if(mainSheet.getRow(row1).getCell(col1).toString().equals((mainSheet.getRow(row2).getCell(col2).toString()))){
			 //System.out.println("yes accurate");
			 return true;
		 }
		 return false;
	 }
	 

}
