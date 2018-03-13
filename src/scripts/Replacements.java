package scripts;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class Replacements {

	public static void IRKeytoCheckmarqKey(int row, int col, XSSFSheet edittedData) {
		String x = edittedData.getRow(row).getCell(2).toString();
		switch (x) {
		case "9.0":
			edittedData.getRow(row).getCell(2).setCellValue(1220);
			break;
		case "10.0":
			edittedData.getRow(row).getCell(2).setCellValue(1200);
			break;
		case "11.0":
			edittedData.getRow(row).getCell(2).setCellValue(1210);
			break;
		case "12.0":
			edittedData.getRow(row).getCell(2).setCellValue(1250);
			break;
		case "13.0":
			edittedData.getRow(row).getCell(2).setCellValue(1230);
			break;
		case "14.0":
			edittedData.getRow(row).getCell(2).setCellValue(1240);
			break;
		case "15.0":
			edittedData.getRow(row).getCell(2).setCellValue(1280);
			break;
		case "16.0":
			edittedData.getRow(row).getCell(2).setCellValue(1260);
			break;
		case "17.0":
			edittedData.getRow(row).getCell(2).setCellValue(1270);
			break;
		case "18.0":
			edittedData.getRow(row).getCell(2).setCellValue(1310);
			break;
		case "19.0":
			edittedData.getRow(row).getCell(2).setCellValue(1290);
			break;
		case "20.0":
			edittedData.getRow(row).getCell(2).setCellValue(1300);
			break;
		case "21.0":
			edittedData.getRow(row).getCell(2).setCellValue(1340);
			break;
		case "22.0":
			edittedData.getRow(row).getCell(2).setCellValue(1320);
			break;
		case "23.0":
			edittedData.getRow(row).getCell(2).setCellValue(1330);
			break;
		case "24.0":
			edittedData.getRow(row).getCell(2).setCellValue(1370);
			break;
		case "25.0":
			edittedData.getRow(row).getCell(2).setCellValue(1350);
			break;
		case "26.0":
			edittedData.getRow(row).getCell(2).setCellValue(1360);
			break;
		case "27.0":
			edittedData.getRow(row).getCell(2).setCellValue(1400);
			break;
		case "28.0":
			edittedData.getRow(row).getCell(2).setCellValue(1380);
			break;
		case "29.0":
			edittedData.getRow(row).getCell(2).setCellValue(1390);
			break;
		case "30.0":
			edittedData.getRow(row).getCell(2).setCellValue(1430);
			break;
		case "31.0":
			edittedData.getRow(row).getCell(2).setCellValue(1410);
			break;
		case "32.0":
			edittedData.getRow(row).getCell(2).setCellValue(1420);
			break;
		case "33.0":
			edittedData.getRow(row).getCell(2).setCellValue(1460);
			break;
		case "34.0":
			edittedData.getRow(row).getCell(2).setCellValue(1440);
			break;
		case "35.0":
			edittedData.getRow(row).getCell(2).setCellValue(1450);
			break;
		case "36.0":
			edittedData.getRow(row).getCell(2).setCellValue(1590);
			break;
		case "37.0":
			edittedData.getRow(row).getCell(2).setCellValue(1470);
			break;
		case "38.0":
			edittedData.getRow(row).getCell(2).setCellValue(1480);
			break;
		case "39.0":
			edittedData.getRow(row).getCell(2).setCellValue(1520);
			break;
		case "40.0":
			edittedData.getRow(row).getCell(2).setCellValue(1500);
			break;
		case "41.0":
			edittedData.getRow(row).getCell(2).setCellValue(1510);
			break;
		case "42.0":
			edittedData.getRow(row).getCell(2).setCellValue(1550);
			break;
		case "43.0":
			edittedData.getRow(row).getCell(2).setCellValue(1530);
			break;
		case "44.0":
			edittedData.getRow(row).getCell(2).setCellValue(1540);
			break;
		case "45.0":
			edittedData.getRow(row).getCell(2).setCellValue(1580);
			break;
		case "46.0":
			edittedData.getRow(row).getCell(2).setCellValue(1560);
			break;
		case "47.0":
			edittedData.getRow(row).getCell(2).setCellValue(1570);
			break;
		case "48.0":
			edittedData.getRow(row).getCell(2).setCellValue(8888);
			break;
		case "49.0":
			edittedData.getRow(row).getCell(2).setCellValue(1590);
			break;
		case "50.0":
			edittedData.getRow(row).getCell(2).setCellValue(1600);
			break;
		case "51.0":
			edittedData.getRow(row).getCell(2).setCellValue(1610);
			break;
		case "52.0":
			edittedData.getRow(row).getCell(2).setCellValue(1620);
			break;
		case "53.0":
			edittedData.getRow(row).getCell(2).setCellValue(1160);
			break;
		case "54.0":
			edittedData.getRow(row).getCell(2).setCellValue(1150);
			break;
		case "55.0":
			edittedData.getRow(row).getCell(2).setCellValue(1190);
			break;
		case "56.0":
			edittedData.getRow(row).getCell(2).setCellValue(1170);
			break;
		case "57.0":
			edittedData.getRow(row).getCell(2).setCellValue(1180);
			break;
		case "58.0":
			edittedData.getRow(row).getCell(2).setCellValue(1630);
			break;
		case "59.0":
			edittedData.getRow(row).getCell(2).setCellValue(1640);
			break;
		case "60.0":
			edittedData.getRow(row).getCell(2).setCellValue(1650);
			break;
		case "61.0":
			edittedData.getRow(row).getCell(2).setCellValue(1660);
			break;
		case "62.0":
			edittedData.getRow(row).getCell(2).setCellValue(1670);
			break;

		}

	}

	public static void convertEvalAuthInits(int row, int col, XSSFSheet edittedData) {
		String x = edittedData.getRow(row).getCell(col).toString();
		switch (x) {
		case "AT":
			edittedData.getRow(row).getCell(col).setCellValue("AThennes");
			break;
		case "B4":
			edittedData.getRow(row).getCell(col).setCellType(Cell.CELL_TYPE_BLANK);
			break;
		case "JB":
			edittedData.getRow(row).getCell(col).setCellValue("JBenjamin");
			break;
		case "JT":
			edittedData.getRow(row).getCell(col).setCellValue("JTrotter");
			break;
		case "KA":
			edittedData.getRow(row).getCell(col).setCellValue("KAtkinson");
			break;
		case "PC":
			edittedData.getRow(row).getCell(col).setCellValue("PCromell");
			break;
		}
	}

	public static void convertEvalNoteItemID(int row, int col, XSSFSheet edittedData) {
		String x = edittedData.getRow(row).getCell(col).toString();
		switch (x) {
		case "26.0":
			edittedData.getRow(row).getCell(col).setCellValue("Dissatisfied");
			break;
		case "27.0":
			edittedData.getRow(row).getCell(col).setCellValue("Satisfied");
			break;
		case "28.0":
			edittedData.getRow(row).getCell(col).setCellValue("Satisfied");
			break;
		case "29.0":
			edittedData.getRow(row).getCell(col).setCellValue("Dissatisfied");
			break;
		case "0.0":
			edittedData.getRow(row).getCell(col).setCellType(Cell.CELL_TYPE_BLANK);
		}

	}

	public void convertIRStudentActivityPlans(int row, XSSFSheet edittedData, int col) {
		String x = edittedData.getRow(row).getCell(col).toString();
		switch (x) {
		case "1.0":
			edittedData.getRow(row).getCell(col).setCellValue("Co-Op");
			break;
		case "2.0":
			edittedData.getRow(row).getCell(col).setCellValue("Internship");
			break;
		case "3.0":
			edittedData.getRow(row).getCell(col).setCellValue("School");
			break;
		case "4.0":
			edittedData.getRow(row).getCell(col).setCellValue("Research");
			break;
		case "5.0":
			edittedData.getRow(row).getCell(col).setCellValue("Flex");
			break;
		case "6.0":
			edittedData.getRow(row).getCell(col).setCellValue("Unconfirmed");
			break;
		case "7.0":
			edittedData.getRow(row).getCell(col).setCellValue("Part-Time-Work");
			break;
		}

	}

}
