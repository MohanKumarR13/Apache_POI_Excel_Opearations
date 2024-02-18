package com.apachepoi.excel.readdatafromprotectedexcel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDataFromProtectedExcel {
	public static void main(String[] args) throws IOException {
		FileInputStream fileInputStream = new FileInputStream(".\\DataFiles\\Customer.xlsx");
		String password = "1234";
		XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(fileInputStream, password);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		int rows = sheet.getLastRowNum();
		System.out.println(rows);
		int column = sheet.getRow(0).getLastCellNum();
		System.out.println(column);
		for (int r = 0; r <= rows; r++) {
			{
				XSSFRow row = sheet.getRow(r);
				for (int c = 0; c < column; c++) {
					XSSFCell cell = row.getCell(c);
					switch (cell.getCellType()) {
					case STRING:
						System.out.print(cell.getStringCellValue());
						break;
					case NUMERIC:
						System.out.print(cell.getNumericCellValue());
						break;
					case BOOLEAN:
						System.out.print(cell.getBooleanCellValue());
						break;
					case FORMULA:
						System.out.print(cell.getNumericCellValue());
						break;
					}
					System.out.print(" ");

				}
				System.out.println();
			}
		}

	}
}
