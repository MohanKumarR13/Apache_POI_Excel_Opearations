package com.apachepoi.excel.operation;

import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcel {

	public static void main(String[] args) throws Exception {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp Info");

		Object empData[][] = { { "EmpID", "Name", "Job" }, { 101, "David", "Engineer" }, { 102, "John", "Manager" },
				{ 103, "Smith", "Analyst" }, };
		// Using for loop
		int rows = empData.length;
		int columns = empData[0].length;

		System.out.println(rows);
		System.out.println(columns);

		for (int row = 0; row < rows; row++) {
			XSSFRow xssfRow = sheet.createRow(row);
			for (int column = 0; column < columns; column++) {
				XSSFCell cell = xssfRow.createCell(column);
				Object value = empData[row][column];
				if (value instanceof String)
					cell.setCellValue((String) value);
				if (value instanceof Integer)
					cell.setCellValue((Integer) value);
				if (value instanceof Boolean)
					cell.setCellValue((Boolean) value);
			}
		}

		// Using for-each loop
		int rowCount = 0;

		for (Object emp[] : empData) {
			XSSFRow row = sheet.createRow(rowCount++);
			int columnNo = 0;
			for (Object value : emp) {
				XSSFCell cell = row.createCell(columnNo++);
				if (value instanceof String)
					cell.setCellValue((String) value);
				if (value instanceof Integer)
					cell.setCellValue((Integer) value);
				if (value instanceof Boolean)
					cell.setCellValue((Boolean) value);
			}
		}

		String filePath = ".\\DataFiles\\employee.xls";
		FileOutputStream outputStream = new FileOutputStream(filePath);
		workbook.write(outputStream);
		outputStream.close();
		System.out.println("Employee.xls file written sucessfully");
	}
}
