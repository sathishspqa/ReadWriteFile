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
		FileInputStream fs = new FileInputStream("C:\\Users\\Light\\eclipse-workspace\\ReadWriteFile\\src\\test\\resources\\InputFile.xlsx");
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
		FileInputStream fs = new FileInputStream("C:\\Users\\Light\\eclipse-workspace\\ReadWriteFile\\src\\test\\resources\\InputFile.xlsx");
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

	public static void main(String[] args) throws IOException {
		
		
		//Path of the excel file
		FileInputStream fs = new FileInputStream("C:\\Users\\Light\\eclipse-workspace\\ReadWriteFile\\src\\test\\resources\\write.xlsx");
		

	 
	        // create blank workbook
	        XSSFWorkbook workbook = new XSSFWorkbook();
	        
	        
	 
	        // Create a blank sheet
	        XSSFSheet sheet = workbook.createSheet("0");
	 
	        ArrayList<Object[]> data = new ArrayList<Object[]>();
	        data.add(new String[] { "ID", "PositionId", "ISIN","QUANTITY","Total Price"});
	        data.add(new Object[] { "1.0", "1.0", "1234.0","3.0","60$" });
	        data.add(new Object[] { "2.0", "2.0", "3456.0","5.0","100$"});
	        data.add(new Object[] { "3.0", "3.0", "9876.0","6.0","30$" });
	       
	 
	        // Iterate over data and write to sheet
	        int rownum = 0;
	        for (Object[] employeeDetails : data) {
	 
	            // Create Row
	            XSSFRow row = sheet.createRow(rownum++);
	 
	            int cellnum = 0;
	            for (Object obj : employeeDetails) {
	 
	                // Create cell
	                XSSFCell cell = row.createCell(cellnum++);
	 
	                // Set value to cell
	                if (obj instanceof String)
	                    cell.setCellValue((String) obj);
	                else if (obj instanceof Double)
	                    cell.setCellValue((Double) obj);
	                else if (obj instanceof Integer)
	                    cell.setCellValue((Integer) obj);
	            }
	        }
	        try {
	 
	            // Write the workbook in file system
	            FileOutputStream out = new FileOutputStream(new File("C:\\Users\\Light\\eclipse-workspace\\ReadExcel\\src\\test\\resources\\write.xlsx"));
	            workbook.write(out);
	            out.close();
	            System.out.println("Data has been created successfully");
	        } catch (Exception e) {
	            e.printStackTrace();
	        } finally {
	            workbook.close();
	        }
	    }
	

}
