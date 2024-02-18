package com.apachepoi.excel.datecells;

import java.io.FileOutputStream;
import java.util.Date;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkingWithDateCells {
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	static XSSFCell cell, cell1, cell1s, s1, s11;
	static CellStyle cellStyle, cellStyle111, cellStyle11, cellStyle1;

	public static void main(String[] args) throws Exception {
		workbook = new XSSFWorkbook();

		sheet = workbook.createSheet("Date format");
		cell = sheet.createRow(0).createCell(0);
		cell.setCellValue(new Date());
		 CreationHelper creationHelper = workbook.getCreationHelper();

		cellStyle = workbook.createCellStyle();
		cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("DD-MM-YYYY"));

		cell1 = sheet.createRow(1).createCell(0);
		cell.setCellValue(new Date());
		cell1.setCellStyle(cellStyle);

		cellStyle1 = workbook.createCellStyle();
		cellStyle1.setDataFormat(creationHelper.createDataFormat().getFormat("MM-DD-YYYY"));

		cell1s = sheet.createRow(2).createCell(0);
		cell1s.setCellValue(new Date());
		cell1s.setCellStyle(cellStyle1);

		cellStyle11 = workbook.createCellStyle();
		cellStyle11.setDataFormat(creationHelper.createDataFormat().getFormat("MM-DD-YYYY hh:mm:ss"));

		s1 = sheet.createRow(3).createCell(0);
		s1.setCellValue(new Date());
		s1.setCellStyle(cellStyle11);

		cellStyle111 = workbook.createCellStyle();
		cellStyle111.setDataFormat(creationHelper.createDataFormat().getFormat("hh:mm:ss"));

		s11 = sheet.createRow(4).createCell(0);
		s11.setCellValue(new Date());
		s11.setCellStyle(cellStyle111);

		FileOutputStream fileOutputStream = new FileOutputStream(".\\DataFiles\\DateFormat.xlsx");

		workbook.write(fileOutputStream);
		workbook.close();
		fileOutputStream.close();
		System.out.println("Done...");
	}
}
