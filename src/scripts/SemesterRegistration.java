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
			int i = 0;
			        for (Row row : mainSheet) {
			            for (Cell cell : row) {
			            	switch (cell.getCellTypeEnum().toString()){
			            	case ("String"):
			            		System.out.println(cell.getRichStringCellValue().getString());
			            	case("NUMERIC")
			            	}
			            	if(cell.getCellTypeEnum().toString() == "STRING"){//janky
				            	System.out.println(cell.getRichStringCellValue().getString());
				            	System.out.println("Row " + row.getRowNum() + " Col: " + cell.getColumnIndex());
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
