package org.whatsup;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Day1 {
	
	public static void main(String[] args) throws IOException    {
		
		File f = new File("C:\\Users\\ARASU\\eclipse-workspace\\FirstDay\\Driver\\book1.xlsx");
		FileInputStream a = new FileInputStream(f);
		Workbook p = new XSSFWorkbook(a);
		Sheet sheet = p.getSheet("DEMO");
		Row row = sheet.getRow(3);
		Cell cell = row.getCell(2);
		String stringCellValue = cell.getStringCellValue();
		System.out.println(stringCellValue);
		
		
	
		
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row2 = sheet.getRow(i);
			
			for (int j = 0; j < row2.getPhysicalNumberOfCells(); j++) {
				Cell cell2 = row2.getCell(j);
				int cellType = cell2.getCellType();
				
				if (cellType ==1) {
					
					String stringCellValue2 = cell2.getStringCellValue();
					System.out.println(stringCellValue2);
				}
				else {
					
					double numericCellValue = cell2.getNumericCellValue();
					long x = (long)numericCellValue;
					System.out.println(x);
					
					
				}
					
					
					
					
					
					
					
				
				
			}
			
		}
		

}}
