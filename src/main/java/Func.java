import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Calendar;
import java.util.regex.Pattern;

import org.apache.poi.hssf.record.RecordInputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//-g
import org.apache.poi.ss.util.CellRangeAddress; 
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
//clasa cu functii necesare
public class Func {

	//determina daca celula este de tip merged, functia va fi folosita la procesarea merged cells
  public boolean isMergedCell(int row, int column) { //Obiectul sheet trebuie sa fie declarat anterior
    for (CellRangeAddress range : sheet.getMergedRegions()) {
        if (range.isInRange(row, column)) {
            return true;
        }
    }

    return false;
}
 //metoda de afisare si stocare orar pe ore
  public static void saveListByHours(FileInputStream fis, int hour) throws IOException {
		Workbook workbook = new XSSFWorkbook(fis);
		//hour = a cata ora [0-6] dintr-o zi
		List<String> rowData = new ArrayList<>();
		int maxColumns = 7*5+4;
		int sumaOreNenule = 0;
		int sumaOreLibere = 0;
		boolean matchingHour = false; 
		try {
			Sheet sheet = workbook.getSheetAt(0);
			for (int i = 5; i < sheet.getPhysicalNumberOfRows(); i++) {
				Row row = sheet.getRow(i);

				// 
				for (int j = 4+hour; j < maxColumns; j=j+7) {
					Cell cell = row.getCell(j);
					String cellValue = (cell == null ? "---" : cell.toString());
					System.out.print(cellValue + "\t");
				}
				//Enter dupa fiecare linie scrisa
				System.out.println();
			}
/* Aici partea de salvare dar care nu functioneaza cum trebuie*/			
/*		for (int j = 4+hour; j < maxColumns; j+=7) { 
			//aici nu mi-a iesit exact formatul de afisare dorit
			//voiam sa afiseze 1 coloana pe fiecare zi
			Row row = sheet.getRow(j);
			Cell cell = row.getCell(j);

			//System.out.print(cell.toString() + ", "); //-g2
			System.out.print("\t");//cate un tab pt fiecare coloana			
			for (int i = 5; i < sheet.getPhysicalNumberOfRows(); i++) {
				row = sheet.getRow(i);
				Cell cell_2 = row.getCell(i);
				if(cell != null)
				{cellValue = cell.getStringCellValue();
				rowData.add(cellValue);
				sumaOreNenule++;}
				else { //sau sa ignoram campurile goale?
					cellValue = "-"; //campurile goale sunt marcate cu -
					sumaOreLibere++;
	//sa facem lista si cu campuri nule, marcate cu "-"?	-->		rowData.add(cellValue);
				}
					
				
				System.out.print(cellValue + "\n");
			}	
		} */
		}finally {
			workbook.close();
			fis.close();}
	}
  
  //Casi, "importCellRangeAddress" nu a fost testat, incearca sa vezi daca merge in main pt un rangeul LuniAnulILicenta
  //creaza o lista de arrays de strings dupa un range ce celule excel, exprimat ca excell de ex: A1:C5
  public static List<String[]> importCellRangeAddress(Sheet sheet, String cellRange) {
      List<String[]> result = new ArrayList<>();

      // Parse the cell range string
      CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf(cellRange);
 
      // Iterate through the rows within the cell range
      for (int rowNum = cellRangeAddress.getFirstRow(); rowNum <= cellRangeAddress.getLastRow(); rowNum++) {
          Row row = sheet.getRow(rowNum);
          if (row == null) {
              // Handle null rows as needed
              continue;
          }

          // Iterate through the cells within the row and cell range
          Iterator<Cell> cellIterator = row.cellIterator();
          List<String> rowData = new ArrayList<>();

          for (int colNum = cellRangeAddress.getFirstColumn(); colNum <= cellRangeAddress.getLastColumn(); colNum++) {
              Cell cell = row.getCell(colNum, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
              // If the cell is null, add an empty string to the rowData list
              rowData.add(cell == null ? "---" : cell.toString());
          }

          // Convert the rowData list to an array of Strings and add it to the result list
          result.add(rowData.toArray(new String[0]));
      }

      return result;
  }

public static boolean isRegexMatchingString(String s, String regx) {
	boolean b=false;
	if(s.isEmpty()||s.equals("---"))
		b = false;
	if(s.matches(regx))
		b = true;
	else {b=false;/*System.out.print(s+" nu indeplineste criteriul de cautare.");*/}
	 return  b;
}

public static boolean isCellMatchingRoomRegex(Cell c) 
{ //cauta Regexul unei celule 
//link to regex > https://regex101.com/r/RzT0oW/2	
	final String regex_intreg = "[A-Z]{2,4}-[SLC]-((PP[1-7])|(PI{1,}[1-7]))-\\s[A-Za-z]*_[A-Z]_?[A-Z]?";
	final String regexDoarSala = "(PP[1-7])|(PI{1,}[1-7])";
	//String regexSalaEtaj = "[A-Z]{2,4}-[SLC]-P(I){1,3}[1-7]-\s[A-Za-z]*_[A-Z]_?[A-Z]?";
	String cellValue = c.getStringCellValue();
	return isRegexMatchingString(cellValue, regex_intreg);
}

}
