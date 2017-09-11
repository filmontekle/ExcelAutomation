package simpleDDT;

import static org.junit.Assert.*;

import java.time.LocalDateTime;
import java.util.concurrent.TimeUnit;

import org.junit.BeforeClass;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import utilities.ExcelUtils;

public class AmazonSearchDDT {
	static WebDriver driver;
	WebElement search;
	WebElement results;

	String excelFilePath = "./src/test/resources/TestData/AmazonSearchData.xlsx";

	@BeforeClass
	public static void setUp() {
		System.setProperty("webdriver.chrome.driver",
				"/Users/filmontekle/Documents/" + "Libraries/drivers/chromedriver");
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		driver.get("https://www.amazon.com");

	}

	@Test
	public void searchTest() throws InterruptedException {

		ExcelUtils.openExcelFile(excelFilePath, "TestData");
		int rowsCount = ExcelUtils.getUsedRowsCount();
		for (int rownum = 1; rownum < rowsCount; rownum++) {
			String searchItem = ExcelUtils.getCellData(rownum, 1);

			searchFor(searchItem);
			String resultText = getSearchResults();
			int resultCount = cleanUpResultsCount(getSearchResults());
			System.out.println("Number of results:" + cleanUpResultsCount(resultText));

			ExcelUtils.setCellData(String.valueOf(resultCount), rownum, 2);

			if (resultCount > 0) {
				System.out.println("pass");
				ExcelUtils.setCellData("pass", rownum, 3);
			} else {
				System.out.println("Fail");
				ExcelUtils.setCellData("Fail", rownum, 3);
			}

			String now = LocalDateTime.now().toString();
			ExcelUtils.setCellData(now, rownum, 4);
		}
	}

	public String getSearchResults() {
		WebDriverWait wait=new WebDriverWait(driver,10);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("s-result-count")));
		
		try {
			results = driver.findElement(By.id("s-result-count"));
		} catch (NoSuchElementException noElem) {
			return "0 results";
		}
		return results.getText();

	}

	public int cleanUpResultsCount(String resultText) {

		String longResult = resultText;
		String[] arrResult = longResult.split(" ");
		int resultsCount;
		if (longResult.contains(" of ")) {

			resultsCount = Integer.parseInt(arrResult[2].replace(",", ""));

		} else {
			resultsCount = Integer.parseInt(arrResult[0].replace(",", ""));
		}
		return resultsCount;

	}
	public void searchFor(String item) throws InterruptedException {
		
		
		search = driver.findElement(By.id("twotabsearchtextbox"));
		search.clear();
		search.sendKeys(item + Keys.ENTER);

	}

}
