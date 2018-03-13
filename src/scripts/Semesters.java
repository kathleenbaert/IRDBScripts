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

public class Semesters {
	
	public static XSSFSheet mainSheet, edittedData;
	public static int edittedDataCurrRow;
	
	public static void main(String [] args) {

		try {
			XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("IR-Source.xlsx"));
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
			FileOutputStream fileOut = new FileOutputStream("Semesters.xlsx");
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
		String [] firstRowEdittedData = new String[] { "ID", "MUID", "TERM", "Activity" };
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
	
	public static void transferLoop(){
		for(int i = 1; i < mainSheet.getPhysicalNumberOfRows(); i++) {
			//CoOpID
			Row row = edittedData.getRow(i);
			if (row == null) {
				row = edittedData.createRow(i);
			}
			edittedData.getRow(i).getCell(0)
			.setCellValue(mainSheet.getRow(i).getCell(0).toString());
			//MUID
			edittedData.getRow(i).getCell(1)
			.setCellValue(mainSheet.getRow(i).getCell(1).toString());
			//Term
			edittedData.getRow(i).getCell(2)
			.setCellValue(mainSheet.getRow(i).getCell(2).toString());
			//Activity
			edittedData.getRow(i).getCell(3)
			.setCellValue(mainSheet.getRow(i).getCell(5).toString());
		}
	}
	public static void cleanUpLoop() {
		for(int i = 0; i < edittedData.getPhysicalNumberOfRows(); i++) {
		Replacements r = new Replacements();
		r.IRKeytoCheckmarqKey(i, edittedData); // works

		r.convertIRStudentActivityPlans(i, edittedData, 3);

		}		
	}
	

}
