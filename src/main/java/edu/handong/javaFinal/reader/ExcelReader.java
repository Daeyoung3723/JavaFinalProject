package edu.handong.javaFinal.reader;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelReader {
	
	public ArrayList<String> getData(InputStream is) {
		ArrayList<String> values = new ArrayList<String>();
		
		try (InputStream inp = is) {
		    //InputStream inp = new FileInputStream("workbook.xlsx");
		    
			XSSFWorkbook workbook = new XSSFWorkbook(inp);
            XSSFSheet sheet = workbook.getSheetAt(0);
            
            int rows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 0; rowIndex < rows; rowIndex++) {
                XSSFRow row = sheet.getRow(rowIndex);
                if (row != null) {
                    int cells = row.getPhysicalNumberOfCells();
                    for (int columnIndex = 0; columnIndex < cells; columnIndex++) {
                        XSSFCell cell = row.getCell(columnIndex);
                        String value = "";
                        if (cell == null)
				            cell = row.createCell(columnIndex);
                        switch (cell.getCellType()) {
                        case NUMERIC:
                            value = cell.getNumericCellValue() + "";
                            break;
                        case STRING:
                            value = cell.getStringCellValue() + "";
                            break;
                        case BLANK:
                        	value = "";
                            break;
                        case _NONE:
                            value = "";
                            break;
                        case ERROR:
                            value = cell.getErrorCellValue() + "";
                            break;
						default:
							break;
                        }
                        values.add(value);
                    }
                }
            }


		        
		    } catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		
		return values;
	}
}
