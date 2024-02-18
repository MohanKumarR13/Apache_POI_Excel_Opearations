package com.apachepoi.excel.datadriventest;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLUtility {
	public FileInputStream fileInputStream;
	public FileOutputStream fileOutputStream;
	public XSSFWorkbook workbook;
	public XSSFSheet sheet;
	public XSSFRow row;
	public XSSFCell cell;
	public CellStyle cellStyle;
	String path = null;

	public XLUtility(String path) {
		this.path = path;
	}

	public int getRowCount(String sheetName) throws Exception {
		fileInputStream = new FileInputStream(path);
		workbook = new XSSFWorkbook(fileInputStream);
		sheet = workbook.getSheet(sheetName);
		int rowCount = sheet.getLastRowNum();
		workbook.close();
		fileInputStream.close();
		return rowCount;
	}

	public int getCellCount(String sheetName, int rowNum) throws Exception {
		fileInputStream = new FileInputStream(path);
		workbook = new XSSFWorkbook(fileInputStream);
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(rowNum);
		int cellCount = row.getLastCellNum();
		workbook.close();
		fileInputStream.close();
		return cellCount;
	}

	public String getCellData(String sheetName, int rowNum, int column) throws Exception {
		fileInputStream = new FileInputStream(path);
		workbook = new XSSFWorkbook(fileInputStream);
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(rowNum);
		cell = row.getCell(column);

		DataFormatter formatter = new DataFormatter();
		String data;
		try {
			data = formatter.formatCellValue(cell);
		} catch (Exception e) {
			data = "";
		}
		workbook.close();
		fileInputStream.close();
		return data;
	}

	public void setCellData(String sheetName, int rowNum, int column, String data) throws Exception {
		fileInputStream = new FileInputStream(path);
		workbook = new XSSFWorkbook(fileInputStream);
		sheet = workbook.getSheet(sheetName);

		row = sheet.getRow(rowNum);
		cell = row.getCell(column);
		cell.setCellValue(data);

		fileOutputStream = new FileOutputStream(path);
		workbook.write(fileOutputStream);
		workbook.close();
		fileInputStream.close();
		fileOutputStream.close();

	}

	public void fillGreenColor(String sheetName, int rowNum, int column) throws Exception {
		fileInputStream = new FileInputStream(path);
		workbook = new XSSFWorkbook(fileInputStream);
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(rowNum);
		cell = row.getCell(column);
		cellStyle = workbook.createCellStyle();
		cellStyle.setFillBackgroundColor(IndexedColors.GREEN.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(cellStyle);
		workbook.write(fileOutputStream);
		workbook.close();
		fileInputStream.close();
		fileInputStream.close();

	}

	public void fillRedColor(String sheetName, int rowNum, int column) throws Exception {
		fileInputStream = new FileInputStream(path);
		workbook = new XSSFWorkbook(fileInputStream);
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(rowNum);
		cell = row.getCell(column);
		cellStyle = workbook.createCellStyle();
		cellStyle.setFillBackgroundColor(IndexedColors.RED.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(cellStyle);
		workbook.write(fileOutputStream);
		workbook.close();
		fileInputStream.close();
		fileOutputStream.close();

	}

	public void setCellData1(String sheetName, int rowNum, int column, String data) throws Exception {
		File xlFile = new File(path);
		if (!xlFile.exists()) {
			workbook = new XSSFWorkbook();
			fileOutputStream = new FileOutputStream(path);
			workbook.write(fileOutputStream);
		}
		fileInputStream = new FileInputStream(path);
		workbook = new XSSFWorkbook(fileInputStream);

		if (workbook.getSheetIndex(sheetName) == 1)
			workbook.createSheet(sheetName);
		sheet = workbook.getSheet(sheetName);

		if (sheet.getRow(rowNum) == null)
			sheet.createRow(rowNum);
		row = sheet.getRow(rowNum);


		cell = row.createCell(column);
		cell.setCellValue(data);
		fileOutputStream = new FileOutputStream(path);
		workbook.write(fileOutputStream);
		workbook.close();
		fileInputStream.close();
		fileOutputStream.close();


	}
}
