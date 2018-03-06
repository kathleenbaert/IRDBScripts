package scripts;

import org.apache.poi.xssf.usermodel.XSSFSheet;

public class boring {
	public static XSSFSheet mainSheet, edittedData;

	public void setMainSheet(XSSFSheet mainSheet) {
		this.mainSheet = mainSheet;
	}

	public void setEdittedData(XSSFSheet edittedData) {
		this.edittedData = edittedData;
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
				// System.out.println("found it in: " + i);
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
				// System.out.println("found it in: " + i);
				return i;
			}
			i++;
		}

	}

	public static int find4992(int row, int col) {
		int i = row;
		while (true) {
			// start search from current row, go down the line
			if (mainSheet.getRow(i).getCell(14).toString().equals("4992.0")) {
				// System.out.println("found it in: " + i);
				return i;
			}
			i++;
		}

	}

	public static int find4994(int row, int col) {
		int i = row;
		while (true) {
			// start search from current row, go down the line
			if (mainSheet.getRow(i).getCell(14).toString().equals("4994.0")) {
				// System.out.println("found it in: " + i);
				return i;
			}
			i++;
		}

	}

	public static boolean verifyMUID(int row1, int col1, int row2, int col2) {
		// System.out.println(mainSheet.getRow(row1).getCell(col1));
		// System.out.println(mainSheet.getRow(row2).getCell(col2));
		if (mainSheet.getRow(row1).getCell(col1).toString().equals((mainSheet.getRow(row2).getCell(col2).toString()))) {
			// System.out.println("yes accurate");
			return true;
		}
		return false;
	}

}
