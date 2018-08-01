package Lecture101CheckErrorCells;

import java.io.FileInputStream;
import java.io.*;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class B {
	
	int setRow;
	String setCell;

	public void reaFiles() throws IOException{
		
		FileInputStream file = new FileInputStream("src/Lecture101CheckErrorCells/Errofile.xlsx");
		XSSFWorkbook book = new XSSFWorkbook(file);
		XSSFSheet sheet = book.getSheet("Tester");
		XSSFRow row;
		XSSFCell cell;
		
		int rowNumber = sheet.getPhysicalNumberOfRows();
		
		for(int i = 0;i<rowNumber;i++) {
			
			row = sheet.getRow(i);
			this.setRow(i); // call method to convert Row to Real Row ID
			int cellNumber = row.getPhysicalNumberOfCells();
			
			for(int j = 0;j<cellNumber;j++) {
				
				cell = row.getCell(j);
				this.setCell(j); // call method to convert cell to Letter ID
				if(row.getCell(j).getCellType()==XSSFCell.CELL_TYPE_STRING) {
					
					if(cell.getStringCellValue().equalsIgnoreCase("Error")) {
						
						System.out.println("Row "+this.setRow(i)+" Cell "+this.setCell(j)+" contain Error");
						
					}
				}
				
			}
			
				
			
		}
		
		
	}

	public int setRow(int setRow) {
		
		this.setRow = setRow+1;
		return this.setRow;
	}
	
	public String setCell(int setCell) {
		
		String cell[] = {"A","B","C","D"};
		
		this.setCell = cell[setCell];
		return this.setCell;
		
	}
	
}
