package com.apachepoi.excel.operation;

import java.io.FileOutputStream;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcel2 {

	public static void main(String[] args) throws Exception {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp Info");
		ArrayList<Object[]> empData = new ArrayList<Object[]>();
		empData.add(new Object[] { "EmpID", "Name", "Job" });
		empData.add(new Object[] { 101, "David", "Engineer" });
		empData.add(new Object[] { 102, "John", "Manager" });
		empData.add(new Object[] { 103, "Smith", "Analyst" });

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
