package com.apachepoi.excel.operation;

import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcel {

	public static void main(String[] args) throws Exception {
		String excelFilePath = ".\\DataFiles\\Countries.xlsx";
		FileInputStream fileInputStream = new FileInputStream(excelFilePath);
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		// XSSFSheet sheet = workbook.getSheet("Sheet1");
		XSSFSheet sheet = workbook.getSheetAt(0);
		// Using For Loop
		int rows = sheet.getLastRowNum();
		int column = sheet.getRow(1).getLastCellNum();
		for (int r = 0; r <= rows; r++) {
			XSSFRow row = sheet.getRow(r);// 0
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

				}
				System.out.print(" ");

			}
			System.out.println();
		}
		// Iterator Method
		Iterator iterator = sheet.iterator();
		while (iterator.hasNext()) {
			XSSFRow row = (XSSFRow) iterator.next();
			Iterator cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				XSSFCell cell = (XSSFCell) cellIterator.next();
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

				}
				System.out.print(" ");

			}
			System.out.println();
		}
	}

}
