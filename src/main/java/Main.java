import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Calendar;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	
	//Stiu ca nu pare mult dar mi-a luat 
	public static void main(String[] args) throws IOException {
		Main rc = new Main();
		//adapteaza in functie de path-ul tau
		File file = new File("D:\\eclipse\\Workspace\\ExcelToCsvConvertor\\orar.xlsx");
		FileInputStream fis = new FileInputStream(file);
		processSchedule(fis);
		//-george-
		Func fx = new Func();
		fx.saveListByHours(fis,0); //vezi functia in clasa Func pt mai multe detalii
		Workbook wb = new XSSFWorkbook(fis);
		CellRangeAddress LuniAnulILicenta = CellRangeAddress.valueOf("E7:K51");
		Sheet thisSheet = wb.getSheetAt(0); // :)
		Sheet newSheet = wb.createSheet("Procesate");
		newSheet.addMergedRegionUnsafe(LuniAnulILicenta);
		//currentsheet.addMergedRegion(LuniAnulILicenta); //-g asta creaza un merge, nu doresc asta, ci sa salveze regiunea intr-un Obiect sheet
		//LuniAnulILicenta.toString();
	}

	private static void processSchedule(FileInputStream fis) throws IOException {
		Workbook workbook = new XSSFWorkbook(fis);
		try {
			Sheet sheet = workbook.getSheetAt(0);

			// Ignora primele 4 coloane(anul,spec,grupa,sgr)
			int startColumn = 4;
			// Incepe prelucrarea in functie de ziua in care esti (momentan doar ce e in paranteza)
			startColumn += (1 - 1) * 7; //-george- nu inteleg de ce te-ai complicat aici, din ce inteleg eu e ca +7 pt ziua urmatoare
			// Seteaza lungimea maxima a 
			int maxColumns = Math.min(sheet.getRow(0).getPhysicalNumberOfCells(),startColumn+ 7);
			// i = 5 pt ca sari peste primele 5 randuri
			for (int i = 5; i < sheet.getPhysicalNumberOfRows(); i++) {
				Row row = sheet.getRow(i);

				// 
				for (int j = startColumn; j < maxColumns; j++) {
					Cell cell = row.getCell(j);
					//String cellValue = (cell != null) ? getCellValue(cell) : ""; 
					String cellValue = (cell != null) ? cell.getStringCellValue() : "---";  
					System.out.print(cellValue + "\t");
			 //-george-ca Obs. am vazut ca nu marcheaza celulele goale cu ---, am incercat sa schimb si conditia if(cell==null), dar tot nu le marcheaza.
			 //-george-cred ca ar trebui descoperita alta metoda in caz ca ne trebuie sa delimitam celulele goale.
				}//Enter dupa fiecare linie scrisa
				System.out.println();
			}
		} finally {
			workbook.close();
			fis.close();
		}

	}
//-george- getCellType e depreciat, propus sa se foloseasca in schimb metoda din pachet getStringCellValue() aplicata unei celule
	//Functie care converteste orice tip de date din excel in string
	private static String getCellValue(Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		case Cell.CELL_TYPE_NUMERIC:
			return Double.toString(cell.getNumericCellValue());
		case Cell.CELL_TYPE_BOOLEAN:
			return Boolean.toString(cell.getBooleanCellValue());
		case Cell.CELL_TYPE_BLANK:
			return "";
		default:
			return "";
		}
	}
}
