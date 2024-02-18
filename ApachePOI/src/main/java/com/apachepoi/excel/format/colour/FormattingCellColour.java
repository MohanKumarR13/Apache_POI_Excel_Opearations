package com.apachepoi.excel.format.colour;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FormattingCellColour {

	public static void main(String[] args) throws Exception {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Sheet1");
		XSSFRow row = sheet.createRow(1);
		// Setting Background Colour
		XSSFCellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setFillBackgroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
		cellStyle.setFillPattern(FillPatternType.BIG_SPOTS);

		XSSFCell cell = row.createCell(1);
		cell.setCellValue("Welcome");
		cell.setCellStyle(cellStyle);

		// Setting foreground color
		cellStyle = workbook.createCellStyle();
		cellStyle.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		cell = row.createCell(2);
		cell.setCellValue("Java");
		cell.setCellStyle(cellStyle);

		FileOutputStream fileOutputStream = new FileOutputStream(".\\DataFiles\\Styles.xlsx");
		workbook.write(fileOutputStream);
		workbook.close();
		fileOutputStream.close();
		System.out.println("Done...");
	}

}
