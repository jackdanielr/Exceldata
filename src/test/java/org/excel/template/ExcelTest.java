package org.excel.template;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Ignore;
import org.testng.annotations.Test;

public class ExcelTest {
	@Ignore
	@Test
	public void test() throws IOException {
		File f = new File(System.getProperty("user.dir") + "/src/test/resources/Feb project 2023.xlsx");
		FileInputStream input = new FileInputStream(f);
		XSSFWorkbook workbook = new XSSFWorkbook(input);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		int totalrow = sheet.getPhysicalNumberOfRows();
		for (int i = 0; i < totalrow; i++) {
			XSSFRow row = sheet.getRow(2);
			int totalcells = row.getPhysicalNumberOfCells();
			for (int j = 0; j < totalcells; j++) {
				XSSFCell cell = row.getCell(2);
				if (cell.getCellType() == CellType.STRING) {
					String s = cell.getStringCellValue();
					System.out.println(s + " ");

				} else {

					double d = cell.getNumericCellValue();
					System.out.println(d + " ");
				}
			}
			System.out.println(" ");

		}
	}



@Test
public void test2() throws IOException {
	File f = new File(System.getProperty("user.dir") + "/src/test/resources/Feb project 2023.xlsx");
	FileInputStream input = new FileInputStream(f);
	XSSFWorkbook workbook = new XSSFWorkbook(input);
	XSSFSheet sheet = workbook.getSheet("Sheet1");
	int totalrow = sheet.getPhysicalNumberOfRows();
//	XSSFRow row=sheet.createRow(18);
//	XSSFCell cell=row.createCell(0);
//	cell.setCellValue(123111);
	
	XSSFRow row=sheet.getRow(16);
	XSSFCell cell=row.getCell(5);
	cell.setCellValue("Helo");
	FileOutputStream output=new FileOutputStream(f);
	workbook.write(output);
	workbook.close();
}
}