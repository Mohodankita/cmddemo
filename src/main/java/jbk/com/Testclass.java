package jbk.com;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class Testclass {
	private static WebDriver driver;
	WebElement we = null;

	List<WebElement> listLinksWebEle = null;
	ArrayList<String> listMenuWeb = null;
	ArrayList<String> listMenuExcel = null;
	Logger log;
	static String pathoflog = "D:\\Testi\\Log4jEx\\demo1\\src\\main\\resources\\log4j.properties";

	@BeforeSuite
	public void openBrowser() {
		driver = new ChromeDriver();
		System.out.println(driver);
		driver.get("file:///D:/Testi/Offline%20Website/pages/examples/dashboard.html");
		System.setProperty("webdriver.chrome.driver", "chromedriver79.exe");
		log = Logger.getLogger(Testclass.class);
		PropertyConfigurator.configure(pathoflog);
		listMenuWeb = new ArrayList<>();
		listMenuExcel = new ArrayList<>();
		we = driver.findElement(By.xpath("/html/body/div[1]/aside/section/ul"));
		listLinksWebEle = we.findElements(By.tagName("li"));
		for (WebElement webElement : listLinksWebEle) {
			listMenuWeb.add(webElement.getText());
		}
		listMenuExcel = getExcelList("D:\\Testi\\Log4jEx\\demo1\\MainLinks.xls", "Mainmenu");

	}

	@Test(dataProvider = "empLogin")
	public void testLinkSpellings(String linkNameExcel) {
		Assert.assertTrue(listMenuWeb.contains(linkNameExcel));
		log.info("testLinkSpellings");
	}

	@Test
	public void testLinksCount() {
		Assert.assertTrue(listMenuWeb.size() == listMenuExcel.size());
		log.info("testLinksCount");
	}

	@Test
	public void testLinksSeq() {
		Assert.assertEquals(listMenuWeb, listMenuExcel);
		if (listMenuWeb.size() == listMenuExcel.size()) {
			for (int i = 0; i < listMenuWeb.size(); i++) {
				boolean flag = true;
				if (!listMenuWeb.get(i).equals(listMenuExcel.get(i))) {
					flag = false;
					break;
				}
				Assert.assertTrue(flag);
			}
		} else {
			Assert.assertTrue(false);
		}
		log.info("testLinksSeq");
	}

	@DataProvider(name = "empLogin")
	public Object[][] loginData() {
		Object[][] arrayObject = getExcelData("D:\\Testi\\Log4jEx\\demo1\\MainLinks.xls", "Mainmenu");
		return arrayObject;
	}

	public ArrayList<String> getExcelList(String fileName, String sheetName) {
		try {
			FileInputStream fs = new FileInputStream(fileName);
			Workbook wb = Workbook.getWorkbook(fs);
			Sheet sh = wb.getSheet("Mainmenu");
			int totalNoOfCols = sh.getColumns();
			int totalNoOfRows = sh.getRows();

			for (int i = 0; i < totalNoOfRows; i++) {
				for (int j = 0; j < totalNoOfCols; j++) {
					listMenuExcel.add(sh.getCell(j, i).getContents());
				}

			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
			e.printStackTrace();
		} catch (BiffException e) {
			e.printStackTrace();
		}
		return listMenuExcel;
	}

	public String[][] getExcelData(String fileName, String sheetName) {
		String[][] arrayExcelData = null;
		try {
			FileInputStream fs = new FileInputStream(fileName);
			Workbook wb = Workbook.getWorkbook(fs);
			Sheet sh = wb.getSheet("Mainmenu");

			int totalNoOfCols = sh.getColumns();
			int totalNoOfRows = sh.getRows();

			arrayExcelData = new String[totalNoOfRows - 1][totalNoOfCols];

			for (int i = 1; i < totalNoOfRows; i++) {

				for (int j = 0; j < totalNoOfCols; j++) {
					arrayExcelData[i - 1][j] = sh.getCell(j, i).getContents();
				}

			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
			e.printStackTrace();
		} catch (BiffException e) {
			e.printStackTrace();
		}
		return arrayExcelData;
	}

}
