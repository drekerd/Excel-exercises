package lecture102ReadDataFromSheetAndDisplayRowAndCell;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class B {

		int row;
		String cell;
		
		public void readFile() throws IOException {
			
			FileInputStream file = new FileInputStream("src/lecture102ReadDataFromSheetAndDisplayRowAndCell/file1.xlsx"); 
			XSSFWorkbook book = new XSSFWorkbook(file);
			XSSFSheet sheet = book.getSheet("Tester");
			XSSFRow row;
			XSSFCell cell;
			int rowTotal = sheet.getPhysicalNumberOfRows();
			
			for(int i = 0; i<rowTotal;i++) {
			
				row = sheet.getRow(i);
				int cellTotal = row.getLastCellNum();
				
				
				for(int j = 0;j<cellTotal;j++) {
					
					cell = row.getCell(j);
					
					this.returnRow(i);
					this.returnCell(j);
					
					
					
					if(j<1) {
						System.out.println("Row "+this.row);
				}
					
					if(cell == null) {
						System.out.println("Cell "+ this.cell+ " Is Empty");
					}else if(row.getCell(j).getCellType()==XSSFCell.CELL_TYPE_NUMERIC) {
						
						System.out.println("Cell " + this.cell +" "+ cell.getNumericCellValue()+"; ");
						
					}else if (row.getCell(j).getCellType()==XSSFCell.CELL_TYPE_STRING) {
						
						System.out.println("Cell "+this.cell+" "+cell.getStringCellValue()+"; ");
						
					}else if (row.getCell(j).getCellType()==XSSFCell.CELL_TYPE_BOOLEAN) {
						
						System.out.println("Cell "+this.cell+" "+cell.getBooleanCellValue()+"; ");
						
					}
				}
				System.out.println();
				
			}
			book.close();
		} //readFile Ends
		
		public void returnRow(int row) {
					this.row = row+1;
		}//returnRow Ends
		
		public void returnCell(int cell) {
			String cellId [] = {"A","B","C","D","E","F","G","H","i","J","K","L","M","N","O","P","Q","R","S","T","V","W","X","Y","Z"};
			
			this.cell = cellId[cell];
		}
} //class ends

		