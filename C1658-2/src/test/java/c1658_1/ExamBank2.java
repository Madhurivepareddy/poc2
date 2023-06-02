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
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.Test;

import net.bytebuddy.implementation.bytecode.Duplication;

public class ExamBank2 extends BaseClass{
	
	@Test(enabled=true)
	public void m2() throws InterruptedException, AWTException, IOException {
         Thread.sleep(5000);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(15));
	
		Actions action = new Actions(driver);
		action.moveToElement(driver.findElement(By.xpath("//span[text()='Add']"))).perform();
		Thread.sleep(1000);
		driver.findElement(By.xpath("//span[contains(text(),'Paper')]")).click();
		driver.findElement(By.xpath("//input[@class='v-textfield v-widget v-textfield-required v-required v-has-width']")).sendKeys(getCellValue("Sheet1", 1, 3));
	
		driver.findElement(By.xpath(
				"//div[@class='v-filterselect v-widget v-filterselect-required v-required v-filterselect-prompt']//div[@role='button']"))
				.click();
		
		Thread.sleep(2000);
		
		Actions act = new Actions(driver);
		act.sendKeys(Keys.chord(Keys.DOWN)).perform();
		driver.findElement(By.xpath("//span[normalize-space()='Reading']")).click();

		WebElement dateWidget = driver.findElement(By.xpath("//button[@type='button']"));
		dateWidget.click();

		Thread.sleep(2000);

		driver.findElement(By.xpath("//span[normalize-space()='18']")).click();
	 //driver.findElement(By.xpath("//body]")).click();
	
	Thread.sleep(3000);
	Robot robot = new Robot();
    for(int i=0; i<3;i++) {
        
        robot.keyPress(KeyEvent.VK_TAB);
        robot.keyRelease(KeyEvent.VK_TAB);
    }
		
		

         //driver.findElement(By.xpath("//input[@class='v-filterselect-input' and @id='gwt-uid-24']")).click();
Thread.sleep(2000);
         WebElement authorTextBox = driver.findElement(By.id("gwt-uid-40"));

		authorTextBox.click();
		authorTextBox.clear();
		authorTextBox.sendKeys("TestAuthor Test");
		driver.findElement(By.xpath("//div[text()='Author']")).click();
		
		Thread.sleep(2000);
		
		
		WebElement addNewbtn = driver.findElement(By.xpath("(//div[@class='v-button v-widget'])[7]"));
        JavascriptExecutor js = (JavascriptExecutor) driver;
        js.executeScript("arguments[0].scrollIntoView(true);", addNewbtn);
        js.executeScript("window.scrollBy(0,250)");
	
		addNewbtn.click();
		Thread.sleep(3000);

		driver.findElement(By.xpath("//span[@class='v-menubar-menuitem v-menubar-menuitem-primary']")).click();
		Thread.sleep(1000);
		
		TakesScreenshot scrShot = ((TakesScreenshot) driver);
		File SrcFile = scrShot.getScreenshotAs(OutputType.FILE);
		File DestFile = new File("C:\\Users\\Admin\\Downloads\\repo_exam-bank-master\\repo_exam-bank-master\\screenshot\\sample2.png");
		FileUtils.copyFile(SrcFile, DestFile);

		driver.findElement(
				By.xpath("//div[@class='v-button v-widget button-9 v-button-button-9 close v-button-close']")).click();
		Thread.sleep(2000);
		driver.findElement(
				By.xpath("//div[@class='v-button v-widget button-1 v-button-button-1 discard v-button-discard']"))
				.click();

		Thread.sleep(1000);
		driver.findElement(By.xpath("//a[normalize-space()='Logout']")).click();
		System.out.println("The screenshot is captured");
		driver.quit();
	}

}
