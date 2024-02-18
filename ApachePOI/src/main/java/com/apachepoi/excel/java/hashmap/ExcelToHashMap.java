package com.apachepoi.excel.java.hashmap;

import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToHashMap {

	public static void main(String[] args) throws Exception {
		FileInputStream fileInputStream = new FileInputStream(".\\DataFiles\\Student.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		XSSFSheet sheet = workbook.getSheet("Student Data");
		int rows = sheet.getLastRowNum();
		HashMap<Object, Object> data = new HashMap<Object, Object>();
		// Reading data from excel to HashMap
		for (int r = 0; r <= rows; r++) {
			String key = sheet.getRow(r).getCell(0).getStringCellValue();
			String value = sheet.getRow(r).getCell(1).getStringCellValue();
			data.put(key, value);
		}
		// Read Data from HashMap
		for (Map.Entry entry : data.entrySet()) {
			System.out.println(entry.getKey() + " " + entry.getValue());
		}

	}

}
