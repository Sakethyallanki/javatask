package ExcelWriting;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Writing1 {
	public static void main(String[] args) throws IOException {

		// it converts into writing mode
		FileOutputStream fos=new FileOutputStream("C:\\Users\\Hello\\eclipse-workspace\\SeleniumP\\EXcelFile\\ExcelWriting.xlsx");
		
		XSSFWorkbook wb=new XSSFWorkbook(); // it is XML spread sheet format
	    XSSFSheet sheet=wb.createSheet(); // it will represents sheets
		
		
	    Scanner sc=new Scanner(System.in);
	    
	    for(int r=0; r<=8; r++) { //it will represents rows -->8
	    	
	    //Create row
	    	
	    		XSSFRow row=sheet.createRow(r);
	    		
	    for(int c=0; c<=5; c++ ) { //it will represents columns-->5
	    	
	    	System.out.println("Enter values");
	    	
	    	String values=sc.next(); // it will accepte string related values
	    	
	    	row.createCell(c).setCellValue(values);	  
	    	
	    }	
	    }
	    
	    wb.write(fos); //writing
	    wb.close(); //close-XSSFWorkbook
	    fos.close();//close-filepath
	
	    System.out.println("Values entering is done");
	}
	

}


