package com.poi.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelFile {
	
	public void readExcel(String filePath, String fileName, String sheetName) throws IOException {
		
		File file = new File(filePath + "\\" + fileName);
		FileInputStream fip = new FileInputStream(file);
		Workbook workbook = null;
		String fileExtensionName = fileName.substring(fileName.indexOf("."));
		
		if(fileExtensionName.equals(".xlsx")) {
			
			workbook = new XSSFWorkbook(fip);
				
		}
		else if(fileExtensionName.equals(".xls")) {
			
			workbook = new HSSFWorkbook(fip);
		}
		
		Sheet sheet = workbook.getSheet(sheetName);
		
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		
		for(int i=0;i<rowCount+1;i++) {
			
			Row row = sheet.getRow(i);
			
			for(int j=0;j<row.getLastCellNum();j++) {
				
				String cell = row.getCell(j).getStringCellValue() + "||";
				System.out.println(cell);
				
			}
			
			System.out.println();
			
		}		
	}
	

	
	public static void main(String[] args) throws IOException {
		
		ReadExcelFile obj = new ReadExcelFile();
		String filePath = System.getProperty("user.dir") + "\\src\\com\\poi\\test";
		obj.readExcel(filePath, "excelFile.xlsx", "excelFileSheet");
		

	}
}
