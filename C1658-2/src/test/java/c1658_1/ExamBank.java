package c1658_1;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.IOException;
import java.time.Duration;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;


public class ExamBank extends BaseClass{

	@BeforeTest
	public void launch() throws InterruptedException, IOException {
		browserLaunch();
		launchUrl(getCellValue("Sheet1", 1, 0));
		maximize();
		driver.findElement(By.id("u")).sendKeys(getCellValue("Sheet1", 1, 1));
		driver.findElement(By.id("p")).sendKeys(getCellValue("Sheet1", 1, 2));
		driver.findElement(By.id("s")).click();


	}

	@Test(enabled=true)
	public void examBank() throws AWTException, InterruptedException, IOException {
		
		Thread.sleep(5000);
		
		Actions action = new Actions(driver);
		action.moveToElement(driver.findElement(By.xpath("//span[text()='Add']"))).perform();
		driver.findElement(By.xpath("//span[text()='Exam']")).click();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(15));

		Thread.sleep(2000);

		Robot robot = new Robot();
		for(int i=0;i<7;i++) {
			
			robot.keyPress(KeyEvent.VK_DOWN);
			robot.keyRelease(KeyEvent.VK_DOWN);
		}

		driver.findElement(
				By.xpath("//div[@class='v-horizontallayout v-layout v-horizontal v-widget']//div[3]//div[1]")).click();
		Thread.sleep(5000);
		
		
			TakesScreenshot screenshot = (TakesScreenshot) driver;
			File srcFile = screenshot.getScreenshotAs(OutputType.FILE);
			
			File destFile = new File("C:\\Users\\Admin\\Downloads\\repo_exam-bank-master\\repo_exam-bank-master\\screenshot\\sample1.png");
			
			FileUtils.copyFile(srcFile, destFile);
	

		driver.findElement(
				By.xpath("//div[@class='v-button v-widget button-9 v-button-button-9 close v-button-close']")).click();
		
		Thread.sleep(1000);

		driver.findElement(
				By.xpath("//div[@class='v-button v-widget button-8 v-button-button-8 cancel v-button-cancel']"))
				.click();
		
		Thread.sleep(1000);

	}
	
	
}
