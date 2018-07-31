package Lecture99_CompareDataFrom2Sheets;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class B {
	
	public void readSheets() throws IOException {
		
		FileInputStream file = new FileInputStream("src/Lecture99_CompareDataFrom2Sheets/fileone.xlsx");
		XSSFWorkbook bookRead = new XSSFWorkbook(file);
		XSSFSheet sheet1 = bookRead.getSheet("Tester");
		
		
		FileInputStream file2 = new FileInputStream("src/Lecture99_CompareDataFrom2Sheets/filetwo.xlsx");
		XSSFWorkbook bookread2 = new XSSFWorkbook(file2);
		XSSFSheet sheet2 = bookread2.getSheet("Tester");
	
		
		int nrRowsFile1 = sheet1.getPhysicalNumberOfRows();
		int nrRowsFile2 = sheet2.getPhysicalNumberOfRows();
		
	
		
		for(int i =0;i<nrRowsFile1;i++) {
			
			XSSFRow row = sheet1.getRow(i);
			XSSFRow row2 = sheet2.getRow(i);
			
			int nrCellsFile1 = row.getPhysicalNumberOfCells();
			int nrCellsFile2 = row.getPhysicalNumberOfCells();
			
			for(int j = 0;j<nrCellsFile1;j++) {
				
				XSSFCell cell = row.getCell(j);
				XSSFCell cell2 = row2.getCell(j);
				if(row.getCell(j).getCellType()==HSSFCell.CELL_TYPE_NUMERIC) {
				
					if(cell.getNumericCellValue()==cell2.getNumericCellValue()) {
						
						System.out.println("Row "+i+" and Cell "+j+" from file 1 is "+row.getCell(j));
						System.out.println("Row "+i+" and Cell "+j+" from file 2 is "+row2.getCell(j));
						System.out.println("Row "+i+" and Cell "+j+" are equal");
					}else {
						System.out.println("Row "+i+" and Cell "+j+" from file 1 is "+row.getCell(j));
						System.out.println("Row "+i+" and Cell "+j+" from file 2 is "+row2.getCell(j));
						System.out.println("Row "+i+" and Cell "+j+" are not equal");
					}
				
					System.out.println();
					
				}
				
				if(row.getCell(j).getCellType()==HSSFCell.CELL_TYPE_STRING) {
					
					if(cell.getStringCellValue().equals(cell2.getStringCellValue())) {
						
						System.out.println("Row "+i+" and Cell "+j+" from file 1 is "+row.getCell(j));
						System.out.println("Row "+i+" and Cell "+j+" from file 2 is "+row2.getCell(j));
						System.out.println("Row "+i+" and Cell "+j+" are equal");
					}else {
						System.out.println("Row "+i+" and Cell "+j+" from file 1 is "+row.getCell(j));
						System.out.println("Row "+i+" and Cell "+j+" from file 2 is "+row2.getCell(j));
						System.out.println("Row "+i+" and Cell "+j+" are not equal");
					}
					System.out.println();
				
				}
				
				
				System.out.println();
				
				
			}
			
			System.out.println();
			
		}
	}
}
		

