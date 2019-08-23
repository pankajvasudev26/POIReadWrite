package com.poi.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelFile {
	
	public void writeExcel(String filePath, String fileName, String sheetName, String[] dataToWrite) throws IOException {
		
		File file = new File(filePath + "\\" + fileName);
		FileInputStream fip = new FileInputStream(file);
		Workbook workbook =null;
		String fileExtensionName = fileName.substring(fileName.indexOf("."));
		
		if(fileExtensionName.equals(".xlsx")) {
			
			workbook = new XSSFWorkbook(fip);
		}
		else if(fileExtensionName.equals(".xls")) {
			
			workbook = new HSSFWorkbook(fip);
		}
		
		Sheet sheet = workbook.getSheet(sheetName);
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		
		Row row = sheet.getRow(0);
		
		Row newRow = sheet.createRow(rowCount + 1);
		
		for(int j=0;j<row.getLastCellNum();j++) {
			
			Cell cell = newRow.createCell(j);
			cell.setCellValue(dataToWrite[j]);
			
		}
		
		fip.close();
		FileOutputStream fos = new FileOutputStream(file);
		workbook.write(fos);
		fos.close();
		
		
	}

	public static void main(String[] args) throws IOException {
		
		String dataToWrite[] = {"Mr. R","Rudransh"};
		
		WriteExcelFile obj = new WriteExcelFile();
		
		String filePath = System.getProperty("user.dir") + "\\src\\com\\poi\\test";
		obj.writeExcel(filePath, "excelFile.xlsx", "excelFileSheet", dataToWrite);
		
		

	}

}
