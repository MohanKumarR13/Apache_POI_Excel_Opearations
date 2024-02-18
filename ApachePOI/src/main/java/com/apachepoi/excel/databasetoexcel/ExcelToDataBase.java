package com.apachepoi.excel.databasetoexcel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToDataBase {
	public static void main(String[] args) throws Exception {
		// Conncet to DataBase
		Connection connection = DriverManager.getConnection("jdbc:mysql://localhost:3306/world", "root", "root");
		// Statement/Query
		Statement statement = connection.createStatement();
		String sql = "create table place(COUNTRYCODE varchar(40),LANGUAGE varchar(40),ISSOFFICIAL varchar(40),PERCENTAGE varchar(40))";
		statement.execute(sql);
		// Excel
		FileInputStream fileInputStream = new FileInputStream(".\\DataFiles\\Country Language.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		XSSFSheet sheet = workbook.getSheet("Country Language");
		int rows = sheet.getLastRowNum();
		for (int r = 1; r <= rows; r++) {
			XSSFRow row = sheet.getRow(r);

			String COUNTRY_CODE = row.getCell(0).getStringCellValue();
			String LANGUAGE = row.getCell(1).getStringCellValue();
			String IS_OFFICIAL = row.getCell(2).getStringCellValue();
			String PERCENTAGE = row.getCell(3).getStringCellValue();

			sql = "insert into place values('" + COUNTRY_CODE + "','" + LANGUAGE + "','" + IS_OFFICIAL + "','"
					+ PERCENTAGE + "')";
			statement.execute(sql);
			statement.execute("commit");
		}
		workbook.close();
		fileInputStream.close();
		connection.close();
		System.out.println("Done...");
	}
}
