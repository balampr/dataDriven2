package org.mo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class modified {
	public static void main(String[] args) throws IOException {
		// locate a file
		File f = new File("F:\\Maven\\DD\\excel\\newExcel.xlsx");
		// read a file
		FileInputStream fi = new FileInputStream(f);
		// workbook - interface
		Workbook book = new XSSFWorkbook(fi);
		// fetch a sheet
		Sheet sh = book.getSheet("data");
		// get a row
		Row row = sh.getRow(0);
		// get a cell
		Cell cell = row.getCell(2);
		//String s = cell.getStringCellValue();
		cell.setCellValue("username");
		
		FileOutputStream fo = new FileOutputStream(f);
		book.write(fo);
		System.out.println("modified.........");
		

	}

}
