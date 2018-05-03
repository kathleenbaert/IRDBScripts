package scripts;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class IR_Student_Plan {

	public static XSSFSheet mainSheet, edittedData;
	public static int edittedDataCurrRow;

	public static void main(String[] args) {

		try {
			// change
			XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("IR-SOURCE3.xlsx"));
			mainSheet = workbook.getSheetAt(0);
			edittedData = workbook.createSheet("edittedData");
			workbook.setMissingCellPolicy(MissingCellPolicy.CREATE_NULL_AS_BLANK);
			workbook.setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

			setUpEdittedData();
			System.out.println("...set up editted data complete");
			transferLoop();
			System.out.println("...transfer loop complete");

			cleanUpLoop();

			System.out.println("...clean up loop complete");
			// write out here

			findDoubles();
			System.out.println("...finding doubles complete");

			FileOutputStream fileOut = new FileOutputStream("IR_Student_Plan.xlsx");
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

	public static void setUpEdittedData() {
		String[] firstRowEdittedData = new String[] { "ID", "MUID", "TERM", "ACTIVITY", "REGISTRATION" };
		for (int i = 0; i < firstRowEdittedData.length; i++) {
			CellReference cr = new CellReference(0, i);
			int r = cr.getRow();
			int c = cr.getCol();
			XSSFRow row1EdittedData = edittedData.getRow(r);
			if (row1EdittedData == null)
				row1EdittedData = edittedData.createRow(r);
			Cell cell = row1EdittedData.getCell(c);
			cell.setCellValue(firstRowEdittedData[i]);
		}
	}

	public static void transferLoop() {
		for (int i = 1; i < mainSheet.getPhysicalNumberOfRows(); i++) {
			// CoOpID
			Row row = edittedData.getRow(i);
			if (row == null) {
				row = edittedData.createRow(i);
			}
			edittedData.getRow(i).getCell(0).setCellValue(mainSheet.getRow(i).getCell(0).toString());
			// MUID
			edittedData.getRow(i).getCell(1).setCellValue(mainSheet.getRow(i).getCell(1).toString());
			// Term
			edittedData.getRow(i).getCell(2).setCellValue(mainSheet.getRow(i).getCell(2).toString());
			// Activity
			edittedData.getRow(i).getCell(3).setCellValue(mainSheet.getRow(i).getCell(5).toString());
			// registration
			edittedData.getRow(i).getCell(4).setCellValue(mainSheet.getRow(i).getCell(14).toString());

		}
	}

	public static void cleanUpLoop() {
		for (int i = 0; i < edittedData.getPhysicalNumberOfRows(); i++) {
			Replacements r = new Replacements();
			r.IRKeytoCheckmarqKey(i, 2, edittedData); // works

			r.convertIRStudentActivityPlans(i, edittedData, 3);
			
			r.convertTermsToNums(i, 4, edittedData);
			
			r.removeNonWorkCompanies(i, edittedData, 5);

		}
	}

	public static void findDoubles() {
		int original = edittedData.getPhysicalNumberOfRows();
		for (int i = 0; i < original - 1; i++) {
			// 8 = registration ID
			if (edittedData.getRow(i).getCell(0).toString().equals(edittedData.getRow(i + 1).getCell(0).toString()) &&

					edittedData.getRow(i).getCell(1).toString().equals(edittedData.getRow(i + 1).getCell(1).toString())

					&& edittedData.getRow(i).getCell(2).toString()
							.equals(edittedData.getRow(i + 1).getCell(2).toString())

					&& edittedData.getRow(i).getCell(3).toString()
							.equals(edittedData.getRow(i + 1).getCell(3).toString())) {// ID, MUID, Term, activity all
																						// equal
				String s1 = edittedData.getRow(i).getCell(4).toString();
				String s2 = edittedData.getRow(i + 1).getCell(4).toString();
				if (s1.equals(s2) || s2.equals("")) {
					edittedData.getRow(i).getCell(4).setCellValue(s1);
				} else {
					edittedData.getRow(i).getCell(4).setCellValue(s1 + ", " + s2);

				}
				s1 = edittedData.getRow(i).getCell(5).toString();
				s2 = edittedData.getRow(i + 1).getCell(5).toString();
				if(s1.equals(s2)) {
					edittedData.getRow(i).getCell(5).setCellValue(s1);
				}
				else if(s1.equals("") && !s2.equals("")) {
					edittedData.getRow(i).getCell(5).setCellValue(s1);
				}else if(s2.equals("") && !s1.equals("")) {
					edittedData.getRow(i).getCell(5).setCellValue(s2);
				}else {
					System.out.println("PROBLEM WITH ID  " + edittedData.getRow(i).getCell(1));
					System.out.println(s1 + "   " + s2);
					
				}
				Row r = edittedData.getRow(i + 1);
				//edittedData.removeRow(r);
				r.getCell(0).setCellValue("BLANK");
				i++;
			}
		}
	}

}