package scripts;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

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
	
	
	
	

}
