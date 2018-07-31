package Lecture_98CopySheetTonNewSheet;
import java.io.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Row;

//import org.apache.poi.xssf.usermodel.XSSFCell;
//import org.apache.poi.xssf.usermodel.XSSFRow;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.xssf.usermodel.*;

public class B {

	
	@SuppressWarnings("deprecation")
	public void readFile() throws IOException{
		
		FileInputStream fs = new FileInputStream("/Users/mariosilva/Documents/Java/TextFilesForTesting/Lecture94Exercise/TestData1.xlsx");
		XSSFWorkbook wk = new XSSFWorkbook(fs);
		XSSFSheet s1 = wk.getSheet("Sheet");
		int r = s1.getPhysicalNumberOfRows();
		
		FileOutputStream os = new FileOutputStream("/Users/mariosilva/Documents/Java/TextFilesForTesting/Lecture94Exercise/FileToWrite/NewFile.xlsx");
		XSSFWorkbook wkbook = new XSSFWorkbook();
		XSSFSheet swrite = wkbook.createSheet("Tester");
		XSSFRow row; //swrite.createRow(0);
		XSSFCell cell; //row.createCell(0);
		
		
		for(int i=0;i<r;i++) {
			
			XSSFRow r1 = s1.getRow(i);
			int c = r1.getPhysicalNumberOfCells();
			row = swrite.createRow(i);
			for(int j = 0;j<c;j++) {
					
					XSSFCell c1 = r1.getCell(j);
					cell = row.createCell(j);
					
					
					if(r1.getCell(j).getCellType()==HSSFCell.CELL_TYPE_NUMERIC) {
						
						cell.setCellValue(c1.getNumericCellValue());
						System.out.print(c1.getNumericCellValue()+" ");
						wkbook.write(os);
						
					}else if(r1.getCell(j).getCellType()==HSSFCell.CELL_TYPE_STRING) {
					
						cell.setCellValue(c1.getStringCellValue());
						
						wkbook.write(os);
						
					}else if(r1.getCell(j).getCellType()==HSSFCell.CELL_TYPE_BOOLEAN){
						cell.setCellValue(c1.getBooleanCellValue());
						System.out.print(c1.getBooleanCellValue()+" ");
							
					}else if(r1.getCell(j).getCellType()==HSSFCell.CELL_TYPE_BLANK) {
						System.out.print("Empty");
					}
					
			}
			
			
		}
		wkbook.write(os);
		wkbook.close();
		wk.close();
		
		
	}

	@SuppressWarnings("deprecation")
	public void sameButWithoutComments() throws IOException {
		
		FileInputStream fs = new FileInputStream("/Users/mariosilva/Documents/Java/TextFilesForTesting/Lecture94Exercise/TestData1.xlsx");
			XSSFWorkbook wk = new XSSFWorkbook(fs);
			XSSFSheet s1 = wk.getSheet("Sheet");
			int r = s1.getPhysicalNumberOfRows();
			
			FileOutputStream os = new FileOutputStream("/Users/mariosilva/Documents/Java/TextFilesForTesting/Lecture94Exercise/FileToWrite/NewFile.xlsx");
			XSSFWorkbook wkbook = new XSSFWorkbook();
			XSSFSheet swrite = wkbook.createSheet("Tester");
			XSSFRow row; //swrite.createRow(0);
			XSSFCell cell; //row.createCell(0);
			
			for(int i=0;i<r;i++) {
				
				XSSFRow r1 = s1.getRow(i);
				int c = r1.getPhysicalNumberOfCells();
				row = swrite.createRow(i);
				for(int j = 0;j<c;j++) {
					
					XSSFCell c1 = r1.getCell(j);
						cell = row.createCell(j);
						
						if(r1.getCell(j).getCellType()==HSSFCell.CELL_TYPE_NUMERIC) {
							cell.setCellValue(c1.getNumericCellValue());							
						}
						
				}
				
				
			}
			wkbook.write(os);
			wkbook.close();
			wk.close();
			
			
		}
		
		
	
	
	public void writeFile()throws IOException{
		
		FileOutputStream file = new FileOutputStream("/Users/mariosilva/Documents/Java/TextFilesForTesting/Lecture94Exercise/FileToWrite/NewFile.xlsx");
		XSSFWorkbook wk = new XSSFWorkbook();
		XSSFSheet sheet = wk.createSheet("Test");
		XSSFRow row; //= sheet.createRow(0);
		XSSFCell cell; //= row.createCell(0);
		
		
		
//		cell.setCellValue("me");
		
		
			row = sheet.createRow(0);
			
			cell = row.createCell(0);
			
//			cell.setCellValue("dada");
		
			cell.setCellValue("cucus");
			wk.write(file);
//			for(int j = 0;j<2;j++) {
//				
//				cell = row.createCell(j);
//				cell.setCellValue("me2");
//				wk.write(file);
//				
//				
//			}
//			
		
		wk.close();
		
		
		
		
		
		
		
	}
	
	
}
