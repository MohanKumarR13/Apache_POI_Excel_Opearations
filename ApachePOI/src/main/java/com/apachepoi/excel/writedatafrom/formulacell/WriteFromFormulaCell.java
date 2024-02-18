package com.apachepoi.excel.writedatafrom.formulacell;

import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteFromFormulaCell {
	public static void main(String[] args) throws Exception {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Numbers");
		XSSFRow row = sheet.createRow(0);

		row.createCell(0).setCellValue(10);
		row.createCell(1).setCellValue(20);
		row.createCell(2).setCellValue(30);

		row.createCell(3).setCellFormula("A1*B1*C1");
		FileOutputStream fileOutputStream = new FileOutputStream(".\\DataFiles\\calc.xlsx");
		workbook.write(fileOutputStream);
		fileOutputStream.close();
		System.out.println("calc.xlsx created with formula cell...");

	}
}
