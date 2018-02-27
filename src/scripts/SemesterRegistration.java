package scripts;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.schemas.office.visio.x2012.main.CellType;

public class SemesterRegistration {

	public static void main(String[] args) {

		try {
			XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("IRAbridged.xlsx"));
			XSSFSheet mainSheet = workbook.getSheetAt(0);
			XSSFSheet edittedData = workbook.createSheet("edittedData");
			// get first row:
			Row row1 = mainSheet.getRow(0);
			String[] firstRow = new String[47];// num of columns across
			for (int i = 0; i < firstRow.length; i++) {
				Cell cell = row1.getCell(i);
				firstRow[i] = cell.getRichStringCellValue().getString();
			}
			// works
			// for(int i = 0; i < firstRow.length; i++){
			// System.out.println(firstRow[i]);
			// }

			// read in rest of the sheet:
			
			int i = 0, j = 0;
			String[][] wholeSheet = new String[mainSheet.getPhysicalNumberOfRows()][47];

			for (Row row : mainSheet) {
	            for (Cell cell : row) {
	            	
	            	switch (cell.getCellTypeEnum().toString()){
	            	case ("STRING"):
	            		wholeSheet[i][j] = cell.getRichStringCellValue().getString();
	            		j++;
	            		break;
	            	case("NUMERIC"):
	            		wholeSheet[i][j] = Double.toString(cell.getNumericCellValue());
	            		j++;
	            		break;
	            	case("BOOLEAN"):
	            		wholeSheet[i][j] = Boolean.toString(cell.getBooleanCellValue());
	            		j++;
	            		break;
	            	default:
	            		wholeSheet[i][j] = "ABORT MISSION!!!";
	            		j++;
	            	}
	            
	            }
	            i++;
	            j = 0;
	        
	    }
//works
		//	System.out.println(Arrays.deepToString(wholeSheet));


			// write out here
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
