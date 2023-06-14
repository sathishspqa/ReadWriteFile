package com.testscenarios;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import com.google.common.collect.Table.Cell;



public class ReadExcel {
	
	@Test (priority = 1)
	
	@SuppressWarnings("incomplete-switch")
	public void readExcel() throws Exception {
		
		
		//Path of the excel file
		FileInputStream fs = new FileInputStream("C:\\Users\\Light\\git\\ReadWriteFile\\ReadWriteFile\\src\\test\\resources\\InputFile.xlsx");
		//Creating a workbook
		XSSFWorkbook workbook = new XSSFWorkbook(fs);
		XSSFSheet sheet = workbook.getSheetAt(0);
		 // Iterate each row one by one
        Iterator<Row> rIterator = sheet.iterator();
        while (rIterator.hasNext()) 
        {
            Row row = rIterator.next();
              
              // For each row, iterate through all the columns
            Iterator<org.apache.poi.ss.usermodel.Cell> Cell = row.cellIterator();
               
            while (Cell.hasNext()) 
            {
                org.apache.poi.ss.usermodel.Cell cell = Cell.next();
                  
                  // Check the cell type
                switch(cell.getCellType())
                {
                case STRING:
                    System.out.print(cell.getStringCellValue());
                    break;
                      
                case NUMERIC:
                    System.out.print(cell.getNumericCellValue()); 
                    break;
                      
                case FORMULA:
                    System.out.print(cell.getNumericCellValue());
                    break;
                }
                System.out.print("|");
            }
            System.out.println();              
        }
        workbook.close();
        fs.close();
		
		
	}
	
	@Test (priority = 2)
	
	public void readExcel2() throws Exception {
		
		
		//Path of the excel file
		FileInputStream fs = new FileInputStream("C:\\Users\\Light\\git\\ReadWriteFile\\ReadWriteFile\\src\\test\\resources\\InputFile.xlsx");
		//Creating a workbook
		XSSFWorkbook workbook = new XSSFWorkbook(fs);
		XSSFSheet sheet = workbook.getSheetAt(1);
		 // Iterate each row one by one
        Iterator<Row> rIterator = sheet.iterator();
        while (rIterator.hasNext()) 
        {
            Row row = rIterator.next();
              
              // For each row, iterate through all the columns
            Iterator<org.apache.poi.ss.usermodel.Cell> Cell = row.cellIterator();
               
            while (Cell.hasNext()) 
            {
                org.apache.poi.ss.usermodel.Cell cell = Cell.next();
                  
                  // Check the cell type
                switch(cell.getCellType())
                {
                case STRING:
                    System.out.print(cell.getStringCellValue());
                    break;
                      
                case NUMERIC:
                    System.out.print(cell.getNumericCellValue()); 
                    break;
                      
                case FORMULA:
                    System.out.print(cell.getNumericCellValue());
                    break;
                }
                System.out.print("|");
            }
            System.out.println();              
        }
        workbook.close();
        fs.close();
		
		
	}
	
	@Test (priority = 3)
	
	public void write() throws Exception {
		
		
		File file = new File("C:\\Users\\Light\\git\\ReadWriteFile\\ReadWriteFile\\src\\test\\resources\\output.xlsx");
		
		 XSSFWorkbook wb = new XSSFWorkbook();
		 
		 XSSFSheet sh = wb.createSheet();
		 
		 sh.createRow(0).createCell(0).setCellValue("ID");
		 sh.getRow(0).createCell(1).setCellValue("PositionId");
		 sh.getRow(0).createCell(2).setCellValue("ISIN");
		 sh.getRow(0).createCell(3).setCellValue("QUANTITY");
		 sh.getRow(0).createCell(4).setCellValue("Total Price");
		 
		 sh.createRow(1).createCell(0).setCellValue(1.0);
		 sh.getRow(1).createCell(1).setCellValue(1.0);
		 sh.getRow(1).createCell(2).setCellValue(1234.0);
		 sh.getRow(1).createCell(3).setCellValue(3.0);
		 sh.getRow(1).createCell(4).setCellValue("60$");
		 
		 sh.createRow(2).createCell(0).setCellValue(2.0);
		 sh.getRow(2).createCell(1).setCellValue(2.0);
		 sh.getRow(2).createCell(2).setCellValue(3456.0);
		 sh.getRow(2).createCell(3).setCellValue(5.0);
		 sh.getRow(2).createCell(4).setCellValue("100$");
		 
		 sh.createRow(3).createCell(0).setCellValue(3.0);
		 sh.getRow(3).createCell(1).setCellValue(3.0);
		 sh.getRow(3).createCell(2).setCellValue(9876.0);
		 sh.getRow(3).createCell(3).setCellValue(6.0);
		 sh.getRow(3).createCell(4).setCellValue("30$");
		 
		 FileOutputStream fos = new FileOutputStream(file);
		 
		 wb.write(fos);
		 
	
		
		
	}
	
	
	

}
