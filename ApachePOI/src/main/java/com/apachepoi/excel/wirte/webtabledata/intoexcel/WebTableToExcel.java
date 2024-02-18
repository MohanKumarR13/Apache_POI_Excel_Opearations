package com.apachepoi.excel.wirte.webtabledata.intoexcel;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.apachepoi.excel.datadriventest.XLUtility;

import io.github.bonigarcia.wdm.WebDriverManager;

public class WebTableToExcel {
	static WebDriver driver;

	@BeforeClass
	public static void main(String[] args) throws Exception {
		WebDriverManager.edgedriver().setup();
		driver = new EdgeDriver();
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		driver.get("https://en.wikipedia.org/wiki/List_of_countries_and_dependencies_by_population");
		String path = ".\\DataFiles\\Population.xlsx";
		XLUtility xlUtility = new XLUtility(path);
		xlUtility.setCellData("Sheet1", 0, 0, "Country");
		xlUtility.setCellData("Sheet1", 0, 1, "Population");
		xlUtility.setCellData("Sheet1", 0, 2, "% of world");
		xlUtility.setCellData("Sheet1", 0, 3, "Date");
		xlUtility.setCellData("Sheet1", 0, 4, "Source");

		WebElement tables = driver
				.findElement(By.xpath("//table[@class='wikitable sortable plainrowheaders jquery-tablesorter']/tbody"));
		int rows = tables.findElements(By.xpath("tr")).size();
		for (int r = 1; r <= rows; r++) {
			String country = tables.findElement(By.xpath("tr[" + r + "]/td[1]")).getText();
			String population = tables.findElement(By.xpath("tr[" + r + "]/td[2]")).getText();
			String perofworld = tables.findElement(By.xpath("tr[" + r + "]/td[3]")).getText();
			String date = tables.findElement(By.xpath("tr[" + r + "]/td[4]")).getText();
			String source = tables.findElement(By.xpath("tr[" + r + "]/td[5]")).getText();

			System.out.println(country + population + perofworld + date + source);
			xlUtility.setCellData("Sheet1", r, 0, country);
			xlUtility.setCellData("Sheet1", r, 1, population);
			xlUtility.setCellData("Sheet1", r, 2, perofworld);
			xlUtility.setCellData("Sheet1", r, 3, date);
			xlUtility.setCellData("Sheet1", r, 4, source);

		}
		System.out.println("Web scraping is done...");
		driver.close();

	}
}
