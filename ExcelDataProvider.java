package assignment5;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataProvider {

	private WebDriver driver = null;
	XSSFWorkbook xwb;
	XSSFSheet sheet1;
	
	@BeforeClass
	public void beforeClass() {
		// System.setProperty("webdriver.chrome.driver",
		// "C://Selenium//chromedriver.exe");
		// driver = new ChromeDriver();

		// System.setProperty("webdriver.gecko.driver",
		// "C://Selenium//geckodriver.exe");
		// driver = new FirefoxDriver();

		// System.setProperty("webdriver.ie.driver",
		// "C://Selenium//IEDriverServer.exe");
		// driver = new InternetExplorerDriver();

		System.setProperty("webdriver.edge.driver", "C://Selenium//MicrosoftWebDriver.exe");
		driver = new EdgeDriver();

		driver.get("http://facebook.com");
		driver.manage().window().maximize();

	}

	@AfterClass
	public void afterClass() throws InterruptedException {
		driver.manage().deleteAllCookies();
		driver.quit();
	}

	@Test(dataProvider = "invalid-login-data")
	public void sampleLoginTest(String username, String password) throws Exception {

		driver.findElement(By.id("email")).sendKeys(username);
		driver.findElement(By.id("pass")).sendKeys(password);
		driver.findElement(By.id("loginbutton")).click();
		Thread.sleep(4000);
		driver.navigate().back();

		driver.findElement(By.id("email")).clear();
	}

	@DataProvider(name = "invalid-login-data")
	public Object[][] dpInavlidLogin() throws Exception {
		
		ExcelDataInput("D://workspacenew1//selenium.xlsx");	
		int rows = getRowCount(0);
		Object[][] data = new Object[rows][2];
		for (int i = 0; i < rows; i++) {
			data[i][0] = getData(0, i, 0);
			data[i][1] = getData(0, i, 1);
			Thread.sleep(1000);
		}
		return data;
	}

	
	public void ExcelDataInput(String Filepath) throws IOException {
	
			File src = new File(Filepath);
			FileInputStream fis = new FileInputStream(src);
			xwb = new XSSFWorkbook(fis);

	}

	public String getData(int SheetNumber, int row, int column) {
		sheet1 = xwb.getSheetAt(SheetNumber);
		String data = sheet1.getRow(row).getCell(column).getStringCellValue();
		return data;
	}

	public int getRowCount(int sheetIndex) {
		int row = xwb.getSheetAt(sheetIndex).getLastRowNum();
		row = row + 1;
		return row;
	}

}
