package testCases;

import java.io.*;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class B{
	
	public void writeFile() throws IOException {
		
		FileOutputStream fs = new FileOutputStream("/Users/mariosilva/Documents/Java/TextFilesForTesting/Lecture93AppacheExcerl/TestData1.xlsx");
		XSSFWorkbook wk = new XSSFWorkbook();
		XSSFSheet s1 = wk.createSheet("Result2");
		XSSFRow r1 = s1.createRow(0);
		XSSFCell c1 = r1.createCell(0);
		
		c1.setCellValue("nizeasdf");
		
		wk.write(fs);
		
		wk.close();	
		
	}
	
	 
	
	
	
}
