/**
 * 
 */
package utilities;

/**
 * @author irfan
 *
 */

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import resources.ExcelConstants;


public class ReadFromExcel {
	FileInputStream fileInputStream;
	Workbook workbook;
	Sheet sheet;
	
	/*
	 * Method to return row count
	 */
	
	
	/*
	 * Method to fetch login data from excel sheet
	 */
	public String[][] getLoginData(String fileName, String sheetName) throws InvalidFormatException, IOException {
		fileName = ExcelConstants.File_Name;
		fileInputStream = new FileInputStream(fileName);
		workbook = WorkbookFactory.create(fileInputStream);
		sheet = workbook.getSheet(ExcelConstants.Sheet_Name);
		int totalNoOfRows = sheet.getPhysicalNumberOfRows();
		System.out.println("Total no. of rows: "+totalNoOfRows);
		int totalCellCount = sheet.getRow(0).getLastCellNum();
		System.out.println("Total cell count : "+totalCellCount);
		
		String loginData[][] = new String[totalNoOfRows][totalCellCount];
		for (int i = 1; i < totalNoOfRows; i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j < totalCellCount; j++) {
				Cell cell = row.getCell(j);
				//System.out.println("cell.getCellType() "+cell.getCellType()+ "   "+cell.getCellTypeEnum());
				try {
					if (cell.getCellType() == cell.getCellTypeEnum()) {
						loginData[i - 1][j] = cell.getStringCellValue();
						System.out.println(" Value present at row number: "+ i + "  at column number " +j+ " : " + loginData[i - 1][j] );
					} 
					else{
						loginData[i - 1][j] = String.valueOf(cell.getNumericCellValue());
						System.out.println(" From else block: Cell value "+loginData[i - 1][j] );
					}
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		}
		return loginData;
	}
	
	public static void main(String args[]) throws InvalidFormatException, IOException{
		ReadFromExcel readexcel = new ReadFromExcel();
		readexcel.getLoginData(ExcelConstants.File_Name,ExcelConstants.Sheet_Name);
	}
}
