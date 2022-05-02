package org.da;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Date;
import java.text.SimpleDateFormat;

import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DRIVEN {
	public static void main(String[] args) throws IOException {
		// locate a file
		File no = new File("F:\\Maven\\DD\\excel\\Excel.xlsx");
		// to read a file
		FileInputStream fin = new FileInputStream(no);
		// workbook - interface
		Workbook bo = new XSSFWorkbook(fin);
		// fetch sheet
		Sheet sh = bo.getSheet("Login");
		// to fetch all data with accurate
		for (int i = 0; i < sh.getPhysicalNumberOfRows(); i++) {
			Row row = sh.getRow(i);

			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);

				int Type = cell.getCellType();

				if (Type == 1) {
					String s = cell.getStringCellValue();
					System.out.println(s);

				} else if (DateUtil.isCellDateFormatted(cell)) {
					java.util.Date date = cell.getDateCellValue();
					System.out.println(date);
					SimpleDateFormat sim = new SimpleDateFormat("dd/MMMMM/yy");
					System.out.println(sim);
				} else {
					double n = cell.getNumericCellValue();
					long l = (long) n;
					System.out.println(l);
					// long ---- string
					String input = String.valueOf(l);
					System.out.println(input);
				}
                
			}
		}

	}

}
