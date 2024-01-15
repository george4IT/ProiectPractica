import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	private static Cell findCell(XSSFSheet sheet, String targetRow, String targetColumn) {
		for (Row row : sheet) {
			Cell firstCellInRow = row.getCell(0);

			// Assuming the first cell in each row contains the row identifier
			if (firstCellInRow != null && firstCellInRow.getCellType() == CellType.STRING) {
				String rowIdentifier = firstCellInRow.getStringCellValue();

				if (rowIdentifier.equals(targetRow)) {
					// Search for the target column in the current row
					for (Cell cell : row) {
						if (cell != null && cell.getCellType() == CellType.STRING) {
							String columnIdentifier = cell.getStringCellValue();

							if (columnIdentifier.equals(targetColumn)) {
								return cell;
							}
						}
					}
				}
			}
		}

		return null; // Cell not found
	}

	public static void main(String[] args) {

		List<String> zile = new ArrayList<String>() {
			{
				add("LUNI");
				add("MARTI");
				add("MIERCURI");
				add("JOI");
				add("VINERI");
			}
		};

		List<String> ore = new ArrayList<String>() {
			{
				add("8-00-9-50");
				add("10-00-11-50");
				add("12-00-13-50");
				add("14-00-15-50");
				add("16-00-17-50");
				add("18-00-19-50");
			}
		};

		List<String> sali = new ArrayList<String>() {
			{
				add("PP1");
				add("PP5");
				add("PP6");
				add("PI1");
				add("PI2");
				add("PII1");
				add("PII2");
				add("PII3");
				add("PII4");
				add("PII5");
				add("PII6");
				add("PII7");
				add("PIII1");
				add("PIII2");
				add("PIII4");
			}
		};

		try {
			String file = "D:/eclipse/Workspace/ExcelToCsvConvertor/orar.xlsx";
			FileInputStream fis = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheetAt(0);

			// Iterator to iterate over rows
			Iterator<Row> itr = sheet.iterator();

			// Skip the first 5 rows
			for (int i = 0; i < 5 && itr.hasNext(); i++) {
				itr.next();
			}

			// Continue iterating over the rest of the rows
			while (itr.hasNext()) {
				Row row = itr.next();
				Iterator<Cell> cellIterator = row.cellIterator();

				// Skip the first 4 columns
				for (int j = 0; j < 4 && cellIterator.hasNext(); j++) {
					cellIterator.next();
				}

				// Iterate over the remaining cells
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();

					switch (cell.getCellType()) {
					case STRING:
						System.out.print(cell.getStringCellValue() + "\t\t\t");
						break;
					case NUMERIC:
						System.out.print(cell.getNumericCellValue() + "\t\t\t");
						break;
					default:
						// Handle other cell types if needed
						break;
					}
				}

				System.out.println("");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
