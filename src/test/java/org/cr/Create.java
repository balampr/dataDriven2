package org.cr;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Create {
	public static void main(String[] args) throws IOException {
		// create a file
		File f = new File("F:\\Maven\\DD\\excel\\newExcel.xlsx");
		// work book - interface
		Workbook book = new XSSFWorkbook();
		// create a sheet
		Sheet sh = book.createSheet("data");
		//create a row
		Row createRow = sh.createRow(0);
		Row Row = sh.createRow(1);
		
		// create a cell
		Cell createCell = createRow.createCell(0);
		Cell createCell2 = createRow.createCell(1);
		Cell createCell3 = createRow.createCell(2);
		// write a value
		createCell.setCellValue("usename");
		createCell2.setCellValue("Password");
		createCell3.setCellValue("Dob");
		// create a cell for row1
		Cell c = Row.createCell(0);
		Cell c1 = Row.createCell(1);
		Cell c2 = Row.createCell(2);
		// write a value for row1
		c.setCellValue("priyanka");
		c1.setCellValue("ahjdcjn");
		c2.setCellValue("28/08/95");
		// write  a sheet
		FileOutputStream fi = new FileOutputStream(f);
		book.write(fi);
		System.out.println("written..............");
		
		
			
	}
}


