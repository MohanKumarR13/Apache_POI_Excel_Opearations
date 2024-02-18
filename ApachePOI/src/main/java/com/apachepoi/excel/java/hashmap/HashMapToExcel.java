package com.apachepoi.excel.java.hashmap;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HashMapToExcel {
	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Student Data");

		Map<Object, Object> data = new HashMap<Object, Object>();
		data.put("101", "John");
		data.put("102", "Smit");
		data.put("102", "Scot");
		data.put("103", "John");
		data.put("104", "Kim");
		data.put("102", "William");
		int rowNo = 0;
		for (Map.Entry entry : data.entrySet()) {
			XSSFRow row = sheet.createRow(rowNo++);

			row.createCell(0).setCellValue((String) entry.getKey());
			row.createCell(1).setCellValue((String) entry.getValue());

		}
		FileOutputStream fileOutputStream = new FileOutputStream(".\\DataFiles\\Student.xlsx");
		workbook.write(fileOutputStream);
		fileOutputStream.close();
		System.out.println("Excel Sheet Created");
	}
}
