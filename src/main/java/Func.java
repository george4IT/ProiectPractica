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
//clasa cu functii necesare
public class Func {

//determina daca celula este de tip merged, functia va fi folosita la procesarea merged cells
protected	CellRangeAddressBase(int firstRow, int lastRow, int firstCol, int lastCol) 
  
  public boolean isMergedCell(int row, int column) {
    for (CellRangeAddress range : sheet.getMergedRegions()) {
        if (range.isInRange(row, column)) {
            return true;
        }
    }

    return false;
}
  //work in progress for searchAfterDayHour
public List<CellRangeAddress> searchAfterDayHour(int column){
  if(sheet.containsColumn(int column))
  return 
    }
  //
  
//CellRangeAddress	getMergedRegion(int index)
//    Returns the merged region at the specified index
//java.util.List<CellRangeAddress>	getMergedRegions()
//    Returns the list of merged regions.
//int	getNumMergedRegions()
//  Returns the number of merged regions


  
/*
setArrayFormula
CellRange<? extends Cell> setArrayFormula(java.lang.String formula,
                                          CellRangeAddress range)
Sets array formula to specified region for result.
Note if there are shared formulas this will invalidate any FormulaEvaluator instances based on this workbook

Parameters:
formula - text representation of the formula
range - Region of array formula for result.
Returns:
the CellRange of cells affected by this change
  */
}
