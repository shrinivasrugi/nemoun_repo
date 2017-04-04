package Day7;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadwrite { 
	
	

	public static void main(String[] args) throws IOException {
		

		//		create input stream
		FileInputStream fis = new FileInputStream("C:\\Users\\DELL\\Desktop\\Ageexp.xlsx");
		
//		create a workbook
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		
//		create sheet
		XSSFSheet sheet = wb.getSheet("Agesheet");
//		XSSFSheet sheet1 = wb.getSheet("Sheet2");
//		get row count
		int rowcount = sheet.getLastRowNum();
//		get column count
		int  colcount = sheet.getRow(0).getLastCellNum();
//		Read values
		
		for(int i =1 ; i<=rowcount;i++){
			
			XSSFCell cell = sheet.getRow(i).getCell(1);
			String celltext = null;
			if(cell.getCellType()==Cell.CELL_TYPE_STRING){
				
				 celltext = cell.getStringCellValue();
			}else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
				celltext = String.valueOf(cell.getNumericCellValue());
				
			}else if(cell.getCellType()==Cell.CELL_TYPE_BLANK){
				
				celltext="";
			}
			
			
			
//			Logic
			double ageval = Double.parseDouble(celltext);
//			write values
			if(ageval > 18){
				
				sheet.getRow(i).getCell(2).setCellValue("Major");
				
			}else{
				sheet.getRow(i).getCell(2).setCellValue("Minor");
			}
		
		}
		
			
//		create output stream
		FileOutputStream fos = new FileOutputStream("C:\\Users\\DELL\\Desktop\\Ageexp.xlsx");
		
//		write and save to excel
		wb.write(fos);
		
//		Close the streams
		fos.close();
		fis.close();
		
		
		
		
		
		
		
		
		
	}

}
