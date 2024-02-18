package com.apachepoi.excel.databasetoexcel;

import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataBaseToExcel {
	public static void main(String[] args) throws Exception {
		// Conncet to DataBase
		Connection connection = DriverManager.getConnection("jdbc:mysql://localhost:3306/world", "root", "root");
		// Statement/Query
		Statement statement = connection.createStatement();
		ResultSet resultSet = statement.executeQuery("Select * From CountryLanguage");
		// Excel
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Country Language");
		XSSFRow row = sheet.createRow(0);
		row.createCell(0).setCellValue("COUNTRYCODE");
		row.createCell(1).setCellValue("LANGUAGE");
		row.createCell(2).setCellValue("ISOFFICIAL");
		row.createCell(3).setCellValue("PERCENTAGE");
		int r = 1;
		while (resultSet.next()) {
			String COUNTRY_CODE = resultSet.getString("COUNTRYCODE");
			String LANGUAGE = resultSet.getString("LANGUAGE");
			String IS_OFFICIAL = resultSet.getString("ISOFFICIAL");
			String PERCENTAGE = resultSet.getString("PERCENTAGE");
			row = sheet.createRow(r++);
			row.createCell(0).setCellValue(COUNTRY_CODE);
			row.createCell(1).setCellValue(LANGUAGE);
			row.createCell(2).setCellValue(IS_OFFICIAL);
			row.createCell(3).setCellValue(PERCENTAGE);

		}
		FileOutputStream fileOutputStream = new FileOutputStream(".\\DataFiles\\Country Language.xlsx");
		workbook.write(fileOutputStream);
		workbook.close();
		fileOutputStream.close();
		System.out.println("Done...");
	}
}
