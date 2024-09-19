package ExcelReading;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reading {
	
public static void main(String[] args) throws IOException {
	//it will convert into reading mode
		FileInputStream fis=new FileInputStream("C:\\Users\\Hello\\eclipse-workspace\\SeleniumP\\EXcelFile\\excel handling.xlsx");
			
			XSSFWorkbook wb=new XSSFWorkbook(fis); // it is XML spread sheet format
		    XSSFSheet sheet=wb.getSheet("Sheet1"); // it will represents sheets
		    
		    //identify rows and columns
		    
		    int rows=sheet.getLastRowNum();
		    int cols=sheet.getRow(1).getLastCellNum();
		    
		    
		    for(int i=0; i<=rows; i++) {// it will represent rows //0,1,2,3,4,5,6,7,8
		    	
		    	XSSFRow crow=sheet.getRow(i);
		    	
		    	for(int c=0; c<cols; c++) { //it will represents columns//0,1,2,3,4,5
		    		
		    		String values =crow.getCell(c).toString();
		    		
		    	System.out.print(values+   "      ");
		    	}
		    	System.out.println();
		    }
		
		}

	
	
}

