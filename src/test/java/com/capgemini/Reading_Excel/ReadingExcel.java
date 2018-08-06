package com.capgemini.Reading_Excel;
import static org.junit.Assert.assertEquals;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;



public class ReadingExcel {

	public static void main(String args[]) throws BiffException, IOException {
		 WebDriver driver;
		 	System.setProperty("webdriver.chrome.driver","D:\\Selenium\\chromedriver.exe");
		String FilePath = "D:\\BDD\\testdataexcel.xls";
		FileInputStream fs = new FileInputStream(FilePath);
		Workbook wb = Workbook.getWorkbook(fs);
		Sheet sh = wb.getSheet("Sheet1");
		int totalNoOfRows = sh.getRows();
		int totalNoOfCols = sh.getColumns();

		for (int row = 0; row < totalNoOfRows; row++) {
			Cell cell[];
			cell=sh.getRow(row);
			driver = new ChromeDriver();
			driver.get("D:/BDD/register.html");
			driver.findElement(By.name("fname")).sendKeys(cell[0].getContents());
			driver.findElement(By.name("lname")).sendKeys(cell[1].getContents());
			driver.findElement(By.name("email")).sendKeys(cell[2].getContents());
			driver.findElement(By.name("mobile")).sendKeys(cell[3].getContents());
			driver.findElement(By.name("address")).sendKeys(cell[4].getContents());
			driver.findElement(By.name("city")).sendKeys(cell[5].getContents());
			Select dropdown= new Select(driver.findElement(By.name("state")));
			dropdown.selectByVisibleText(cell[6].getContents());
			driver.findElement(By.name("submit")).click();
			driver.get("D:/BDD/projectdetails.html");
			driver.findElement(By.name("pname")).sendKeys(cell[7].getContents());
			driver.findElement(By.name("cname")).sendKeys(cell[8].getContents());
			driver.findElement(By.name("tsize")).sendKeys(cell[9].getContents());
			driver.findElement(By.name("submit")).click();
			Alert alert = driver.switchTo().alert();
			 System.out.println(alert.getText());
			 assertEquals("Registered Successfully",alert.getText());
			driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
			
		}
	

	
		
	}

}