package scripts;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.schemas.office.visio.x2012.main.CellType;


public class SemesterRegistration {
	
	public static void main (String [] args){
		
		
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("IRAbridged.xlsx"));
			XSSFSheet mainSheet = workbook.getSheetAt(0);
			XSSFSheet edittedData = workbook.createSheet("edittedData");
			//do what needs to be done
			//get first row:
			Row row1 = mainSheet.getRow(0);
			String [] firstRow = new String [47];
			for (int i = 0; i < firstRow.length; i++){
				Cell cell = row1.getCell(i);
				firstRow[i] = cell.getRichStringCellValue().getString();
			}
			//works
//			for(int i = 0; i < firstRow.length; i++){
//				System.out.println(firstRow[i]);
//			}
			System.out.println("\n\n\n");
			for (Row row : mainSheet) {
			            for (Cell cell : row) {
			            	System.out.println(cell.getCellTypeEnum().toString());
			            	
			            	switch (cell.getCellTypeEnum().toString()){
			            	case ("STRING"):
			            		System.out.println(cell.getRichStringCellValue().getString());
			            		break;
			            	case("NUMERIC"):
			            		System.out.println(cell.getNumericCellValue());
			            		break;
			            	case("BOOLEAN"):
			            		System.out.println(cell.getBooleanCellValue());
			            		break;
			            	default:
			            		System.out.println("THERE WAS A PROBLEM!!!!!");
			            	}
			            
			            }
			        
			    }
			
			
			
			
			//write out here
			FileOutputStream fileOut = new FileOutputStream("DataOut.xlsx");
			workbook.write(fileOut);
			fileOut.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}

}
