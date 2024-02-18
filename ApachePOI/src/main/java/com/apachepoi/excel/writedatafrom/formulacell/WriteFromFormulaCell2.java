package com.apachepoi.excel.writedatafrom.formulacell;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteFromFormulaCell2 {
	public static void main(String[] args) throws Exception {

		String path = ".\\DataFiles\\Books.xlsx";

		FileInputStream fileInputStream = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		//XSSFSheet sheet = workbook.getSheet("Sheet1");
		XSSFSheet sheet = workbook.getSheetAt(0);

		sheet.getRow(7).getCell(2).setCellFormula("SUM(C2:C6)");
		fileInputStream.close();

		FileOutputStream fileOutputStream = new FileOutputStream(path);
		workbook.write(fileOutputStream);
		workbook.close();
		fileOutputStream.close();
		System.out.println("Done...");

	}
}
