package testCases;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.usermodel.*;

public class A {

	public static void main(String[] args) throws IOException {
		
		
		A read = new A();
		B write = new B();
		
		
		
//		read.readFromFile();
		
//		read.findHowManyRows();
//		read.findHowManyCells();
		write.writeFile();
//		read.readEntireFile();
	
	}
	
	//Read From File Exemple1
	public void readFromFile() throws IOException{
		
		FileInputStream fs = new FileInputStream("/Users/mariosilva/Documents/Java/TextFilesForTesting/Lecture93AppacheExcerl/TestData1.xlsx");
		

		
		// Create Workbook Class Object
		
		//XSLS - Class name will start from XSSF, and in case of XLS, Class name start by HSSF
		
		XSSFWorkbook wk = new XSSFWorkbook(fs);
		
		XSSFSheet s1 = wk.getSheet("Plan1");
		XSSFRow r1 = s1.getRow(0);
		XSSFCell c1 = r1.getCell(0);
		
		System.out.println(c1.getStringCellValue());
		
		
	}

	
	//Find How Many Rows are in the file
	public void findHowManyRows() throws IOException{
		
		FileInputStream fs = new FileInputStream("/Users/mariosilva/Documents/Java/TextFilesForTesting/Lecture93AppacheExcerl/TestData1.xlsx");
		
		// Create Workbook Class Object
		
		//XSLS - Class name will start from XSSF, and in case of XLS, Class name start by HSSF
		
		XSSFWorkbook wk = new XSSFWorkbook(fs);
		
		XSSFSheet s1 = wk.getSheet("Plan1");
		
		System.out.println("Number of Rows "+s1.getPhysicalNumberOfRows()); // this line will return how many rows are filled in the sheet
		System.out.println("Number Of Last Row "+s1.getLastRowNum()); // lets you know the last row that is being used
		
	}

	public void findHowManyCells() throws IOException {
		
		
		FileInputStream fs = new FileInputStream("/Users/mariosilva/Documents/Java/TextFilesForTesting/Lecture93AppacheExcerl/TestData1.xlsx");
		
		// Create Workbook Class Object
		
		//XSLS - Class name will start from XSSF, and in case of XLS, Class name start by HSSF
		
		XSSFWorkbook wk = new XSSFWorkbook(fs);
		XSSFSheet s1 = wk.getSheet("Plan1");
		XSSFRow r1 = s1.getRow(8);
		
		System.out.println("Number of Cells "+r1.getPhysicalNumberOfCells());
		System.out.println("Number of Last cells Used "+r1.getLastCellNum());
	
		
		
	}


	public void readEntireFile() throws IOException{
		
		FileInputStream fs = new FileInputStream("/Users/mariosilva/Documents/Java/TextFilesForTesting/Lecture93AppacheExcerl/TestData2.xlsx");
		
		XSSFWorkbook wk = new XSSFWorkbook(fs);
		XSSFSheet s1 = wk.getSheet("Plan1");
//		XSSFRow r1 = s1.getRow(0);
		
		int r = s1.getPhysicalNumberOfRows();
//		
		
//		System.out.println(r);
//		System.out.println(c);
		int j = 0;
		int i = 0;
		
		for(i = 0;i<r;i++) {
			
//			System.out.println(s1.getRow(i));
			XSSFRow r1 = s1.getRow(i);
			int c = r1.getPhysicalNumberOfCells();
			
				for(j =0;j<c;j++) {
			
					
					System.out.print(r1.getCell(j));
//					XSSFCell c1 = r1.getCell(j);
//					System.out.print(c1.getStringCellValue());
					System.out.print(" ");
				}
			
			System.out.println("\n");
			
		}
		
		
		
	}
	

}
