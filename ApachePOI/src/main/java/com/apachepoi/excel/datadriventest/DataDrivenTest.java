package com.apachepoi.excel.datadriventest;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class DataDrivenTest {
	public static WebDriver driver;

	@BeforeClass
	public void setUp() {
		WebDriverManager.edgedriver().setup();
		driver = new EdgeDriver();
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		driver.manage().window().maximize();
	}

	@Test(dataProvider = "LoginData")
	public void loginTest(String user, String pwd, String exp) {
		driver.get("https://admin-demo.nopcommerce.com/login");
		WebElement email = driver.findElement(By.id("Email"));
		email.clear();
		email.sendKeys(user);

		WebElement pwds = driver.findElement(By.id("Password"));
		pwds.clear();
		pwds.sendKeys(pwd);

		driver.findElement(By.xpath("//button[normalize-space()='Log in']")).click();

		String exp_title = "Dashboard / nopCommerce administration";
		String act_title = driver.getTitle();

		if (exp.equals("Valid")) {
			if (exp_title.equals(act_title)) {
				driver.findElement(By.xpath("//*[@id=\"navbarText\"]/ul/li[3]/a")).click();
				Assert.assertTrue(true);
			} else {
				Assert.assertTrue(false);

			}
		} else if (exp.equals("Invalid")) {
			if (exp_title.equals(act_title)) {
				driver.findElement(By.xpath("//*[@id=\"navbarText\"]/ul/li[3]/a")).click();
				Assert.assertTrue(false);
			} else {
				Assert.assertTrue(true);

			}
		}
	}

	@DataProvider(name = "LoginData")
	public String[][] getData() throws Exception {
	/*	String loginData[][] = { { "admin@yourstore.com", "admin", "Valid" },
				{ "admin@yourstore.com", "adm", "Invalid" }, { "adm@yourstore.com", "admin", "Invalid" },
				{ "adm@yourstore.com", "adm", "Invalid" } }; */
		
		String path=".\\DataFiles\\loginData.xlsx";
		XLUtility xlUtility=new XLUtility(path);
		int totalRows=xlUtility.getRowCount("Sheet1");
		int totalCoumns=xlUtility.getCellCount("Sheet1", 1);
		
		String loginData[][]=new String[totalRows][totalCoumns];
		for(int i=1;i<=totalRows;i++) {
			for(int j=0;j<totalCoumns;j++) {
				loginData[i-1][j]=xlUtility.getCellData("Sheet1", i, j);
			}
		}
		return loginData;
	}

	@AfterClass
	void tearDown() {
		driver.close();
	}
}
