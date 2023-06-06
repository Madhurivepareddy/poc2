package com.TestSapCrm.mavenproject;



	import java.io.File;
	import java.io.FileInputStream;
	import java.io.FileNotFoundException;
	import java.io.FileOutputStream;
	import java.io.IOException;
	import java.io.InterruptedIOException;
	import java.time.Duration;
	import java.util.Date;
	import java.util.Iterator;
	import java.util.LinkedHashMap;
	import java.util.List;
	import java.util.Map;
	import java.util.Set;

	import org.apache.commons.io.FileUtils;
	import org.apache.commons.lang3.ObjectUtils.Null;
	import org.apache.hc.core5.util.Timeout;
	import org.apache.poi.xssf.usermodel.XSSFCell;
	import org.apache.poi.xssf.usermodel.XSSFRow;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;
	import org.apache.xmlbeans.XmlException;
	import org.openqa.selenium.By;
	import org.openqa.selenium.JavascriptExecutor;
	import org.openqa.selenium.Keys;
	import org.openqa.selenium.NoSuchElementException;
	import org.openqa.selenium.OutputType;
	import org.openqa.selenium.TakesScreenshot;
	import org.openqa.selenium.WebDriver;
	import org.openqa.selenium.WebElement;
	import org.openqa.selenium.chrome.ChromeDriver;
	import org.openqa.selenium.chrome.ChromeOptions;
	import org.openqa.selenium.devtools.v109.indexeddb.model.Key;
	import org.openqa.selenium.interactions.Action;
	import org.openqa.selenium.interactions.Actions;

	import org.openqa.selenium.support.ui.ExpectedConditions;
	import org.openqa.selenium.support.ui.Select;
	import org.openqa.selenium.support.ui.WebDriverWait;
	import org.openqa.selenium.support.events.EventFiringWebDriver;
	import org.testng.Assert;
	import org.testng.annotations.AfterClass;
	import org.testng.annotations.BeforeClass;
	// import org.junit.Test;
	import org.testng.annotations.BeforeTest;
	import org.testng.annotations.Test;
	import org.testng.asserts.SoftAssert;

	import io.github.bonigarcia.wdm.WebDriverManager;

	/**
	 * Unit test for simple App.
	 */
	public class CrmMultiline {
	    WebDriver driver = null;
	    WebElement Username;
	    WebElement Password;
	    WebElement Logon;
	    WebElement SideArrow;
	    WebElement SalesClk;
	    WebElement agssales;
	    WebElement Crm;
	    WebElement Saprm;
	    WebElement ContractManager;
	    WebElement ele;
	    WebElement newclick;
	    WebElement ProductIdv;
	    String LgnTittle;
	    String ActualTittle;
	    String Actualdate = "06.06.2022";
	    String Price1 = "1.000,00";
	    String Price2 = "2.000,00";
	    String Price3 = "3.000,00";
	    WebElement Dropdown1;
	    WebElement itemlist;
	    WebElement ExpandContractMaint;
	    SoftAssert softAssert = new SoftAssert();
	    String textnum;
	    WebElement scroll1;
	    String MaintananceNO;
	    String ServiceNo;
	    WebElement Milstone1;
	    WebElement curbill;
	    WebElement orgbill;
	    WebElement verifyorignalbil;

	    WebElement UI5;
	    WebElement ProjectMaint;
	    FileInputStream fis;
	    XSSFWorkbook wb;
	    XSSFSheet sheet1;
	    File file;
	    //ScreenShot
	    String[] ScreenshotNames= new String[100];
		int array_increment=0;
	  ReusableScreenshot reuse;
	  XSSFWorkbook workbook;
		// Declare An Excel Work Sheet
		XSSFSheet sheet;
		// Declare A Map Object To Hold TestNG Results
		Map<String, Object[]> TestNGResults;
		public static String driverPath = "C:\\Users\\X0143782\\OneDrive - Applied Materials\\Documents\\ScreenShot";
	    

	    @BeforeTest
	    public void LaunchChromeBrowser() throws IOException {
	        WebDriverManager.chromedriver().setup();
	        // Open Chrome Browser
	        ChromeOptions options = new ChromeOptions();
	        options.addArguments("--remote-allow-origins=*");
	        driver = new ChromeDriver(options);
	        driver.manage().window().maximize();

	        driver.get("http://sap-db2qacloudtest/");
	    }

	    @Test(priority = 1)

	    public void loginpage() throws InterruptedException, IOException {
	        File file = new File("C:\\Users\\X0143782\\Downloads\\exceldata1.xlsx");
	        FileInputStream fis = new FileInputStream(file);

	        XSSFWorkbook wb = new XSSFWorkbook(fis);
	        XSSFSheet sheet1 = wb.getSheetAt(0);

	        String username = sheet1.getRow(0).getCell(1).getStringCellValue();
	        System.out.println("user name is: " + username);

	        String password = sheet1.getRow(1).getCell(1).getStringCellValue();
	        System.out.println("password is " + password);
	        driver.findElement(By.xpath("//*[@id='logonuidfield']")).sendKeys(username);
	        driver.findElement(By.xpath("//*[@id='logonpassfield']")).sendKeys(password);

	        Thread.sleep(4000);// C:\Users\X0143783\OneDrive - Applied Materials\Documents\CRM
	                           // app\democrm\src\test\java\com\crmproject\Screenshot
	                           reuse= new ReusableScreenshot();
	                           reuse.captureScreenshot(ScreenshotNames[array_increment++]="Login Page", driver);
	        Logon = driver.findElement(By.name("uidPasswordLogon"));
	        Logon.click();
	        TestNGResults.put("2", new Object[] { 1d, "Login to SapCrm", "Login Successful", "Pass" });
	        System.out.println("Login Successful");
	    }

	    @Test(priority = 2)
	    public void verifyLoginpage() {

	        LgnTittle = driver.getTitle();
	        ActualTittle = "SAP NetWeaver Portal";
	        Assert.assertEquals(LgnTittle, ActualTittle);
	        TestNGResults.put("2", new Object[] { 2d, "Checking Home Page Tittle", "Home Page Tittle-SAP NetWeaver Portal", "Pass" });

	    }

	    @Test(priority = 3)
	    public void EnterprisePortal() throws InterruptedException {
	        Thread.sleep(6000);

	        // WebElement SideArrow;
	        // SideArrow= wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(
	        // "//*[@id='tlnOverflowBtn']")));
	        // SideArrow.click();

	        SideArrow = driver.findElement(By.xpath("//*[@id='tlnOverflowBtn']"));
	        SideArrow.click();

	        SalesClk = driver
	                .findElement(By.xpath("//*[@id='topTabMenuItem33']/td[2]/div"));
	        SalesClk.click();
	    }

	    @Test(priority = 4)
	    public void agspage() throws InterruptedException {

	        Thread.sleep(5000);

	        agssales = driver.findElement(By.xpath("//*[@id='subTabIndex1']/div[1]"));

	        agssales.click();

	        Thread.sleep(4000);

	        Crm = driver.findElement(By.xpath("//*[@id='L2N1']"));
	        Crm.click();
	        Set<String> handle = driver.getWindowHandles(); // Switch to new chrome window
	        Iterator<String> it = handle.iterator();
	        String Parentwindow = it.next();
	        String Childwindow = it.next();
	        System.out.println(Parentwindow);

	        driver.switchTo().window(Childwindow);

	        System.out.println(Childwindow);
	        Thread.sleep(10000);

	        String l = driver.getTitle();

	        System.out.println(l);

	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='contentAreaFrame']")));

	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='isolatedWorkArea']")));
	        driver.findElement(By.linkText("ZCOM_SUPUSR-Contract Manager")).click();
	        Thread.sleep(15000);
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='CRMApplicationFrame']")));

	        // driver.switchTo().frame(driver.findElement(By.xpath("//frame[@id='HeaderFrame']")));

	        driver.switchTo().frame(driver.findElement(By.xpath("//frame[@id ='WorkAreaFrame1']")));
	        Thread.sleep(7000);

	        driver.findElement(By.xpath("//*[@id='C4_W16_V17_SRV-CONTR']")).click();

	        Thread.sleep(2000);
	        driver.findElement(By.cssSelector("a[id='C24_W69_V70_SRV-SCO-SR']")).click();
	    }

	    @Test(priority = 5)
	    public void newpage() throws InterruptedException,XmlException, IOException {
	        Thread.sleep(2000);
	        driver.findElement(By.xpath("//*[@id='C26_W76_V77_V79_thtmlb_button_1']")).click();
	        Thread.sleep(5000);
	        Set<String> handle = driver.getWindowHandles(); // Switch to new chrome window
	        Iterator<String> it = handle.iterator();
	        String Parentwindow = it.next();
	        String Childwindow = it.next();
	        String subchild = it.next();
	        System.out.println(Parentwindow);

	        driver.switchTo().window(subchild);

	        System.out.println(Childwindow);
	        System.out.println(subchild);

	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='WorkAreaFrame1popup']")));
	        Thread.sleep(3000);

	        driver.findElement(By.xpath("//*[@id='C28_W85_V86_proctype_table[2].process_type']")).click();
	        Thread.sleep(2000);
	        driver.switchTo().window(Childwindow);
	        System.out.println(Childwindow);
	        Thread.sleep(10000);
	    }

	    @Test(priority = 6)
	    public void mulwindow() throws InterruptedException, IOException {
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='contentAreaFrame']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='isolatedWorkArea']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='CRMApplicationFrame']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//frame[@id ='WorkAreaFrame1']")));
	        Thread.sleep(2000);

	        File file = new File("C:\\Users\\X0143782\\Downloads\\exceldata1.xlsx");
	        FileInputStream fis = new FileInputStream(file);

	        XSSFWorkbook wb = new XSSFWorkbook(fis);
	        XSSFSheet sheet1 = wb.getSheetAt(0);
	        int soldtoparty = (int) sheet1.getRow(2).getCell(1).getNumericCellValue();
	        String i = String.valueOf(soldtoparty);
	        System.out.println("entered data is " + soldtoparty);
	        ele = driver.findElement(By.cssSelector("input[id*='_btpartnerset_soldto_name']"));
	        ele.sendKeys(i);
	        Thread.sleep(3000);
	        driver.findElement(By.xpath("//*[@id='C29_W88_V91_V93_btdatecontractstart_date']")).sendKeys("01.04.2022");
	        driver.findElement(By.xpath("//*[@id='C29_W88_V91_V93_btdatecontractend_date']")).sendKeys("01.04.2023");
	        Thread.sleep(5000);
	        ele.sendKeys(Keys.ENTER);
	        Thread.sleep(50000);
	        Set<String> handle = driver.getWindowHandles(); // Switch to new chrome window
	        Iterator<String> it = handle.iterator();
	        String Parentwindow = it.next();
	        String Childwindow = it.next();
	        String subchild = it.next();
	        // String sub = it.next();
	        driver.switchTo().window(subchild);
	        // driver.switchTo().window(Childwindow);

	        Thread.sleep(2000);
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='WorkAreaFrame1popup']")));
	        // click on 3058 scenerio
	        driver.findElement(By.cssSelector("span[id*='_btqrsorgdet_table[1].description_orgunit_1']")).click();
	        Thread.sleep(5000);
	        driver.switchTo().window(Childwindow);
	    }

	    @Test(priority = 7)
	    public void shiptoparty() throws InterruptedException,XmlException, IOException {
	        Thread.sleep(7000);
	        Set<String> handle = driver.getWindowHandles(); // Switch to new chrome window
	        Iterator<String> it = handle.iterator();
	        String Parentwindow = it.next();
	        String Childwindow = it.next();
	        String subchild = it.next();
	        driver.switchTo().window(subchild);
	        // click on ship to party button
	        System.out.println(subchild);
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='WorkAreaFrame1popup']")));
	        driver.findElement(By.cssSelector("div[id*='_btqrspartdet_table[1].default_partner__css']")).click();
	        Thread.sleep(4000);
	        driver.findElement(By.cssSelector("div[id*='_btqrspartdet_table[1].default_partner__css']")).click();
	        driver.switchTo().window(Childwindow);
	        Thread.sleep(1000);

	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='contentAreaFrame']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='isolatedWorkArea']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='CRMApplicationFrame']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//frame[@id ='WorkAreaFrame1']")));
	        JavascriptExecutor js = (JavascriptExecutor) driver;
	        js.executeScript("window.scrollBy(0,400)", "");

	        Thread.sleep(2000);
	        WebElement apgmulti = driver.findElement(By.cssSelector("input[id*='_btadminh_po_number_sold']"));
	        apgmulti.sendKeys("APG multi line");
	        apgmulti.sendKeys(Keys.ENTER);
	        reuse= new ReusableScreenshot();
	        reuse.captureScreenshot(ScreenshotNames[array_increment++]="SoldToPartyInput", driver);
	        Thread.sleep(3000);

	        driver.findElement(By.xpath("//*[@id='C29_W88_V91_0002_expand_link']")).click();
	        driver.findElement(By.xpath("//*[@id='C29_W88_V91_0001_expand_link']")).click();
	        driver.findElement(By.xpath("//*[text()='Organizational Data']")).click();
	        Thread.sleep(2000);
	        driver.findElement(By.cssSelector("a[id*='_btorgset_struct.service_org_short-btn']")).click();
	    }

	    @Test(priority = 8)
	    public void organisationdata() throws InterruptedException, IOException {
	        Thread.sleep(4000);
	        Set<String> handle = driver.getWindowHandles(); // Switch to new chrome window
	        Iterator<String> it = handle.iterator();
	        String Parentwindow = it.next();
	        String Childwindow = it.next();
	        String subchild = it.next();
	        driver.switchTo().window(subchild);
	        driver.switchTo().frame(driver.findElement(By.xpath("//*[@id='f4modalframe']")));
	        Dropdown1 = driver.findElement(By.xpath("//*[text()='AMAT_UK_SP']"));
	        Actions actions = new Actions(driver);
	        actions.scrollToElement(Dropdown1).perform();
	        Thread.sleep(5000);
	        Dropdown1.click();
	        driver.switchTo().window(Childwindow);
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='contentAreaFrame']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='isolatedWorkArea']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='CRMApplicationFrame']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//frame[@id ='WorkAreaFrame1']")));

	        driver.findElement(By.xpath("//*[@id='C29_W88_V91_0012_expand_link']")).click();
	        driver.findElement(By.xpath("//*[@id='C29_W88_V91_0003_expand_link']")).click();
	        Thread.sleep(7000);

	        itemlist = driver.findElement(By.xpath("//*[text()='Items']"));

	        Thread.sleep(3000);
	        itemlist.click();

	        Thread.sleep(5000);
	        

	        // item list 1st
	        File file = new File("C:\\Users\\X0143782\\Downloads\\exceldata1.xlsx");
	        FileInputStream fis = new FileInputStream(file);

	        XSSFWorkbook wb = new XSSFWorkbook(fis);
	        XSSFSheet sheet1 = wb.getSheetAt(0);
	        String itemlist1 = sheet1.getRow(3).getCell(1).getStringCellValue();
	        System.out.println("entered data is " + itemlist1);
	        WebElement Product1 = driver.findElement(By.cssSelector("input[id*='_btadmini_table[1].ordered_prod']"));
	        Product1.click();
	        Product1.sendKeys(itemlist1);
	        WebElement P1 = driver.findElement(By.cssSelector("input[id*='_btadmini_table[1].quantity']"));
	        P1.click();
	        Thread.sleep(2000);
	        P1.sendKeys("1");
	        P1.sendKeys(Keys.ENTER);

	        Thread.sleep(5000);
	        // item list 2nd

	        String itemlist2 = sheet1.getRow(4).getCell(1).getStringCellValue();
	        System.out.println("entered data is " + itemlist2);

	        WebElement Product2 = driver.findElement(By.cssSelector("input[id*='_btadmini_table[2].ordered_prod']"));
	        Product2.click();
	        Product2.sendKeys(itemlist2);
	        WebElement P2 = driver.findElement(By.cssSelector("input[id*='_btadmini_table[2].quantity']"));
	        P2.click();
	        Thread.sleep(2000);
	        P2.sendKeys("1");
	        P2.sendKeys(Keys.ENTER);
	        Thread.sleep(4000);
	        // item list 3rd
	        String itemlist3 = sheet1.getRow(5).getCell(1).getStringCellValue();
	        System.out.println("entered data is " + itemlist3);

	        WebElement Product3 = driver.findElement(By.cssSelector("input[id*='_btadmini_table[3].ordered_prod']"));
	        Product3.click();
	        Product3.sendKeys(itemlist3);
	        WebElement P3 = driver.findElement(By.cssSelector("input[id*='_btadmini_table[3].quantity']"));
	        P3.click();
	        Thread.sleep(2000);
	        P3.sendKeys("1");
	        P3.sendKeys(Keys.ENTER);
	        Thread.sleep(5000);
	        // drop down check
	        driver.findElement(By.cssSelector("button[id*='_btadmini_table[1].itm_type-btn']")).click();
	        driver.findElement(By.partialLinkText("NonTlCt TimeRR-CO")).click();
	        Thread.sleep(3000);
	        driver.findElement(By.cssSelector("button[id*='_btadmini_table[2].itm_type-btn']")).click();
	        driver.findElement(By.partialLinkText("NBC WBS SrvRR-CO")).click();
	        Thread.sleep(2000);
	        driver.findElement(By.cssSelector("button[id*='_btadmini_table[3].itm_type-btn']")).click();
	        driver.findElement(By.partialLinkText("NBC Servics CtC POC")).click();
	        Thread.sleep(3000);
	        driver.findElement(By.cssSelector("button[id*='_btadmini_table[2].zzafld000006-btn']")).click();
	        driver.findElement(By.partialLinkText("License"))
	                .click();
	        Thread.sleep(2000);
	        driver.findElement(By.cssSelector("img[id*='_btadmini_table[1].thtmlb_oca.EDIT']")).click();
	        Thread.sleep(5000);
	        driver.findElement(By.xpath("//*[text()='Price Details']")).click(); // click on price detail
	        Thread.sleep(4000);
	        driver.findElement(By.xpath("//*[text()='Add']")).click(); // click on add
	        Thread.sleep(4000);
	        reuse= new ReusableScreenshot();
	        reuse.captureScreenshot(ScreenshotNames[array_increment++]="ItemList", driver);
	    }

	    @Test(priority = 9)

	    public void pricedetail() throws InterruptedException, IOException {
	        reuse= new ReusableScreenshot();

	        Set<String> handle = driver.getWindowHandles(); // Switch to new chrome window
	        Iterator<String> it = handle.iterator();
	        String Parentwindow = it.next();
	        String Childwindow = it.next();
	        String subchild = it.next();
	        driver.switchTo().window(subchild);
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='WorkAreaFrame1popup']")));
	        driver.findElement(By.xpath("//*[text()='AMAT: Fix price/tool']")).click();
	        Thread.sleep(2000);
	        driver.switchTo().window(Childwindow);

	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='contentAreaFrame']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='isolatedWorkArea']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='CRMApplicationFrame']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//frame[@id ='WorkAreaFrame1']")));
	        File file = new File("C:\\Users\\X0143782\\Downloads\\exceldata1.xlsx");
	        FileInputStream fis = new FileInputStream(file);

	        XSSFWorkbook wb = new XSSFWorkbook(fis);
	        XSSFSheet sheet1 = wb.getSheetAt(0);
	        int price = (int) sheet1.getRow(6).getCell(1).getNumericCellValue();
	        String i1 = String.valueOf(price);
	        System.out.println("entered data is " + price);
	        WebElement price1 = driver.findElement(By.cssSelector("input[id*='condline_table[1].kbetr_prt']"));
	        Thread.sleep(3000);
	        price1.clear();
	        price1.sendKeys(i1);
	        price1.sendKeys(Keys.ENTER);

	        Thread.sleep(3000);

	        driver.findElement(By.xpath("//*[text() ='Billing Plan']")).click(); // click on bill plan
	        Thread.sleep(2000);
	        driver.findElement(By.cssSelector("a[id*='C41_W168_V169_V173_thtmlb_button_1']")).click(); // click on insert
	        Thread.sleep(4000);

	        WebElement pric1 = driver.findElement(By.cssSelector("input[id*='_btbillplandet_table[1].bill_date']"));
	        pric1.click();
	        pric1.sendKeys("06.06.2022");

	        WebElement P2 = driver.findElement(By.cssSelector("input[id*='_btbillplandet_table[1].bill_value']"));
	        P2.click();
	        P2.sendKeys("1000");
	        P2.sendKeys(Keys.ENTER);
	        Thread.sleep(4000);

	        driver.findElement(By.cssSelector("a[title='Next Item']")).click(); // click on next for line item 20
	        Thread.sleep(4000);
	        driver.findElement(By.xpath("//*[text()='Add']")).click(); // click on add
	        Thread.sleep(4000);
	    }

	    @Test(priority = 10)
	    public void line20() throws InterruptedException, IOException {
	        reuse= new ReusableScreenshot();
	        Set<String> handle = driver.getWindowHandles(); // Switch to new chrome window
	        Iterator<String> it = handle.iterator();
	        String Parentwindow = it.next();
	        String Childwindow = it.next();
	        String subchild = it.next();

	        driver.switchTo().window(subchild);
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='WorkAreaFrame1popup']")));
	        Thread.sleep(2000);
	        driver.findElement(By.xpath("//*[text()='AMAT: Var Price/Tool']")).click();
	        Thread.sleep(2000);
	        driver.switchTo().window(Childwindow);
	        Thread.sleep(2000);
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='contentAreaFrame']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='isolatedWorkArea']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='CRMApplicationFrame']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//frame[@id ='WorkAreaFrame1']")));
	        Thread.sleep(2000);
	        File file = new File("C:\\Users\\X0143782\\Downloads\\exceldata1.xlsx");
	        FileInputStream fis = new FileInputStream(file);
	        XSSFWorkbook wb = new XSSFWorkbook(fis);
	        XSSFSheet sheet1 = wb.getSheetAt(0);
	        int price = (int) sheet1.getRow(7).getCell(1).getNumericCellValue();
	        String i1 = String.valueOf(price);
	        System.out.println("entered data is " + price);
	        WebElement price123 = driver.findElement(By.cssSelector("input[id*='condline_table[1].kbetr_prt']"));
	        price123.clear();
	        price123.sendKeys(i1);
	        price123.sendKeys(Keys.ENTER);

	        Thread.sleep(4000);

	        driver.findElement(By.cssSelector("a[id*='C41_W168_V169_V173_thtmlb_button_1']")).click(); // click on insert
	        Thread.sleep(4000);

	        WebElement price2 = driver.findElement(By.cssSelector("input[id*='_btbillplandet_table[1].bill_date']"));
	        price2.click();
	        price2.sendKeys("06.06.2023");

	        WebElement P21 = driver.findElement(By.cssSelector("input[id*='_btbillplandet_table[1].bill_value']"));
	        P21.click();
	        P21.sendKeys("2000");
	        P21.sendKeys(Keys.ENTER);
	        Thread.sleep(4000);

	        WebElement Val1 = driver.findElement(By.cssSelector("a[id*='_btbillplandet_table[2].hyperlink']"));
	        if (Val1.isDisplayed()) {
	            System.out.println("Element for Line item 20 is Displayed");
	            reuse.captureScreenshot(ScreenshotNames[array_increment++]="LineItem20", driver);
	            TestNGResults.put("3", new Object[] { 1d, "Enter Price For Line Item 20", "Pod Is Enabled", "Pass" });
	        } else {
	            System.out.println("Element not present");
	            TestNGResults.put("3", new Object[] { 1d, "Enter Price For Line Item 20", "Pod Is Enabled", "Fail" });
	        }

	        driver.findElement(By.cssSelector("a[title='Next Item']")).click(); // click on next for line item 30
	        Thread.sleep(4000);
	        driver.findElement(By.xpath("//*[text()='Add']")).click(); // click on add
	        Thread.sleep(4000);
	    }

	    @Test(priority = 11)
	    public void lineitem30() throws InterruptedException, IOException {
	        reuse= new ReusableScreenshot();
	        Set<String> handle = driver.getWindowHandles(); // Switch to new chrome window
	        Iterator<String> it = handle.iterator();
	        String Parentwindow = it.next();
	        String Childwindow = it.next();
	        String subchild = it.next();
	        driver.switchTo().window(subchild);
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='WorkAreaFrame1popup']")));
	        Thread.sleep(4000);
	        driver.findElement(By.xpath("//*[text()='AMAT: Var Price/Tool']")).click();
	        Thread.sleep(2000);
	        driver.switchTo().window(Childwindow);
	        Thread.sleep(2000);
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='contentAreaFrame']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='isolatedWorkArea']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='CRMApplicationFrame']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//frame[@id ='WorkAreaFrame1']")));
	        Thread.sleep(2000);
	        File file = new File("C:\\Users\\X0143782\\Downloads\\exceldata1.xlsx");
	        FileInputStream fis = new FileInputStream(file);
	        XSSFWorkbook wb = new XSSFWorkbook(fis);
	        XSSFSheet sheet1 = wb.getSheetAt(0);
	        int price = (int) sheet1.getRow(8).getCell(1).getNumericCellValue();
	        String i1 = String.valueOf(price);
	        System.out.println("entered data is " + price);

	        WebElement price3 = driver.findElement(By.cssSelector("input[id*='condline_table[1].kbetr_prt']"));
	        price3.clear();
	        price3.sendKeys(i1);
	        price3.sendKeys(Keys.ENTER);

	        Thread.sleep(4000);

	        driver.findElement(By.cssSelector("a[id*='C41_W168_V169_V173_thtmlb_button_1']")).click(); // click on insert
	        Thread.sleep(4000);

	        driver.findElement(By.cssSelector("input[id*='_btbillplandet_table[1].bill_date']"))
	                .sendKeys("06.06.2022");

	        WebElement P31 = driver.findElement(By.cssSelector("input[id*='_btbillplandet_table[1].bill_value']"));
	        P31.click();
	        P31.sendKeys("3000");
	        P31.sendKeys(Keys.ENTER);
	        Thread.sleep(4000);

	        driver.findElement(By.xpath("//*[text()='Back']")).click(); // back button
	        Thread.sleep(4000);
	        driver.findElement(By.xpath("//*[text()='Save']")).click(); // save
	        Thread.sleep(2000);
	        reuse.captureScreenshot(ScreenshotNames[array_increment++]="ProfitCentre", driver);


	        WebElement Profit1 = driver.findElement(By.cssSelector("span[id*='_btadmini_table[1].zzbfld000001']"));
	        if (Profit1.isDisplayed()) {
	            String Procntr1 = Profit1.getText();
	            System.out.println("Profit Centre For Line Item 10 Is Generated " + Procntr1);
	            TestNGResults.put("4", new Object[] { 1d, "Checcking ProfitCenter for Line Item 10 ", "Profit Center Created", "Pass" });

	        } else {
	            System.out.println("ProfitCentre is not Displayed for Line Item 10");
	            TestNGResults.put("4", new Object[] { 1d, "Checcking ProfitCenter for Line Item 10 ", "Profit Center Created", "Fail" });

	        }

	        Thread.sleep(4000);
	        WebElement Profit12 = driver.findElement(By.cssSelector("span[id*='_btadmini_table[1].zzbfld000001']"));
	        if (Profit12.isDisplayed()) {
	            String Procntr12 = Profit12.getText();
	            System.out.println("Profit Centre For Line Item 20 Is Generated " + Procntr12);
	            TestNGResults.put("5", new Object[] { 1d, "Checcking ProfitCenter for Line Item 20 ", "Profit Center Created", "Pass" });

	        } else {
	            System.out.println("ProfitCentre is not Displayed for Line Item 20");
	            TestNGResults.put("5", new Object[] { 1d, "Checcking ProfitCenter for Line Item 10 ", "Profit Center Created", "Fail" });
	        }

	        WebElement Profit3 = driver.findElement(By.cssSelector("span[id*='_btadmini_table[1].zzbfld000001']"));
	        if (Profit3.isDisplayed()) {
	            String Procntr3 = Profit3.getText();
	            System.out.println("Profit Centre For Line Item 30 Is Generated " + Procntr3);
	            TestNGResults.put("6", new Object[] { 1d, "Checcking ProfitCenter for Line Item 30 ", "Profit Center Created", "Pass" });

	        } else {
	            System.out.println("ProfitCentre is not Displayed for Line Item 30");
	            TestNGResults.put("6", new Object[] { 1d, "Checcking ProfitCenter for Line Item 30 ", "Profit Center Created", "Fail" });
	            Thread.sleep(2000);
	        }

	        // Thread.sleep(3000);
	        // driver.findElement(By.xpath("a[id*='_0003_expand_link']")).click();

	        Thread.sleep(2000);
	        driver.findElement(By.xpath("//*[text()='Bill Plan']")).click();
	        Thread.sleep(20000);

	        driver.findElement(By.xpath("//*[text()='PO Maintenance']")).click();
	        // driver.findElement(By.xpath("//*[@id=C29_W88_V91_0029_expand_link]")).click();

	    }

	    @Test(priority = 12)
	    public void VerifyUnderPoMaintanance() throws InterruptedException,XmlException, IOException {
	        reuse= new ReusableScreenshot();

	        Thread.sleep(4000);
	        Set<String> handle = driver.getWindowHandles(); // Switch to new chrome window
	        Iterator<String> it = handle.iterator();
	        String Parentwindow = it.next();
	        String Childwindow = it.next();
	        String subchild = it.next();
	        driver.switchTo().window(subchild);
	        Thread.sleep(4000);
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='WorkAreaFrame1popup']")));
	        Thread.sleep(4000);
	        reuse.captureScreenshot(ScreenshotNames[array_increment++]="PO Maintainance", driver);

	        WebElement Dateline10 = driver.findElement(By.cssSelector("span[id*='_pomaintain_table[1].bill_date']"));
	        String Date10 = Dateline10.getText();
	        softAssert.assertEquals(Date10, Actualdate);
	        System.out.println("Date For Line Item 10 UnderPo maintanance is " + Date10);
	        Thread.sleep(2000);

	        WebElement Dateline20 = driver.findElement(By.cssSelector("span[id*='_pomaintain_table[4].bill_date']"));
	        String Date20 = Dateline20.getText();
	        softAssert.assertEquals(Date20, Actualdate);
	        System.out.println("Date For Line Item 20 UnderPo maintanance is " + Date20);
	        Thread.sleep(2000);

	        WebElement Dateline30 = driver.findElement(By.cssSelector("span[id*='_pomaintain_table[2].bill_date']"));
	        String Date30 = Dateline30.getText();
	        softAssert.assertEquals(Date30, Actualdate);
	        System.out.println("Date For Line Item 30 UnderPo maintanance is " + Date30);
	        Thread.sleep(2000);

	        WebElement BillPrice10 = driver
	                .findElement(By.cssSelector("span[id*='_pomaintain_table[1].bill_value']"));
	        String BillPr10 = BillPrice10.getText();
	        softAssert.assertEquals(BillPr10, Price1);
	        System.out.println("Price For Line Item 10 UnderPo maintanance is " + BillPr10);
	        TestNGResults.put("7", new Object[] { 1d, "Checking Billprice10 under PO maintainance ", "PO is Created", "Pass" });
	        Thread.sleep(2000);

	        WebElement BillPrice20 = driver
	                .findElement(By.cssSelector("span[id*='_pomaintain_table[4].bill_value']"));
	        String BillPr20 = BillPrice20.getText();
	        softAssert.assertEquals(BillPr20, Price2);
	        System.out.println("Price For Line Item 20 UnderPo maintanance is " + BillPr20);
	        TestNGResults.put("8", new Object[] { 1d, "Checking Billprice20 under PO maintainance ", "PO is Created", "Pass" });
	        Thread.sleep(2000);

	        WebElement BillPrice30 = driver
	                .findElement(By.cssSelector("span[id*='_pomaintain_table[2].bill_value']"));
	        String BillPr30 = BillPrice30.getText();
	        softAssert.assertEquals(BillPr30, Price3);
	        System.out.println("Price For Line Item 30 UnderPo maintanance is " + BillPr30);
	        TestNGResults.put("9", new Object[] { 1d, "Checking Billprice30 under PO maintainance ", "PO is Created", "Pass" });
	        Thread.sleep(4000);
	        driver.findElement(By.xpath("//*[text()='CANCEL']")).click(); // click on Cancel
	        Thread.sleep(6000);
	        driver.switchTo().window(Childwindow);
	        System.out.println(Childwindow);
	        Thread.sleep(10000);
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='contentAreaFrame']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='isolatedWorkArea']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='CRMApplicationFrame']")));
	        driver.switchTo().frame(driver.findElement(By.xpath("//frame[@id ='WorkAreaFrame1']")));

	        driver.findElement(By.xpath("//*[text()='Bill Plan']")).click();
	        Thread.sleep(3000);

	        ExpandContractMaint = driver.findElement(By.xpath("//*[text()='Contract Maintenance']"));

	        Actions actions = new Actions(driver);

	        actions.scrollToElement(ExpandContractMaint).perform();

	        Thread.sleep(4000);
	        ExpandContractMaint.click();

	    }

	    @Test(priority = 13)
	    public void contractmaintanance10() throws InterruptedException, IOException {
	        Thread.sleep(2000);
	        reuse= new ReusableScreenshot();

	        WebElement line10 = driver
	                .findElement(By.cssSelector("input[id*= 'searchquerynode_parameters[2].VALUE1']"));
	        line10.sendKeys("10");

	        driver.findElement(By.cssSelector("input[id*= 'searchquerynode_parameters[3].VALUE1']"))
	                .sendKeys("01.04.2022");
	        driver.findElement(By.cssSelector("input[id*= 'searchquerynode_parameters[4].VALUE1']"))
	                .sendKeys("01.04.2023");

	        line10.sendKeys(Keys.ENTER);
	        reuse.captureScreenshot(ScreenshotNames[array_increment++]="ontractmaintanance10", driver);
	        Thread.sleep(4000);

	        WebElement Bill10 = driver.findElement(By.cssSelector("span[id*='_resultnode_table[1].billing_date']"));
	        String Bill1 = Bill10.getText();
	        WebElement Price10 = driver
	                .findElement(By.cssSelector("span[id*='_resultnode_table[1].billing_amount']"));
	        String Price11 = Price10.getText();
	        softAssert.assertEquals(Bill1, Actualdate);
	        System.out.println("Billing Date For Line Item 10 " + "=" + Bill1);
	        TestNGResults.put("10", new Object[] { 1d, "Check billing price", "bill price displayed for line item 10", "Pass" });
	        softAssert.assertEquals(Price11, Price1);

	        System.out.println("Billing Price For Line Item 10 " + "=" + Price11);

	        Thread.sleep(2000);
	        File file = new File("C:\\Users\\X0143782\\Downloads\\exceldata1.xlsx");
	        FileInputStream fis = new FileInputStream(file);

	        XSSFWorkbook wb = new XSSFWorkbook(fis);
	        XSSFSheet sheet1 = wb.getSheetAt(0);
	        int bill = (int) sheet1.getRow(9).getCell(1).getNumericCellValue();
	        String i1 = String.valueOf(bill);
	        System.out.println("entered data is " + bill);
	        WebElement line30 = driver.findElement(By.cssSelector("input[id*= 'searchquerynode_parameters[2].VALUE1']"));
	        line30.clear();
	        line30.sendKeys(i1);
	        line30.sendKeys(Keys.ENTER);
	        Thread.sleep(4000);

	        WebElement Bill30 = driver.findElement(By.cssSelector("span[id*='_resultnode_table[1].billing_date']"));
	        String Bill3 = Bill30.getText();
	        WebElement Price30 = driver
	                .findElement(By.cssSelector("span[id*='_resultnode_table[1].billing_amount']"));
	        String Price31 = Price30.getText();
	        softAssert.assertEquals(Bill3, Actualdate);
	        System.out.println("Billing Date For Line Item 30 " + "=" + Bill3);
	        TestNGResults.put("11", new Object[] { 1d, "Check billing price", "bill price displayed for line item 30", "Pass" });
	        softAssert.assertEquals(Price31, Price3);

	        System.out.println("Billing Price For Line Item 30 " + "=" + Price31);
	        Thread.sleep(3000);

	        WebElement Items = driver.findElement(By.cssSelector("a[id*='_0003_expand_link']"));
	        Actions actionss = new Actions(driver);
	        actionss.scrollToElement(Items).perform();
	        Thread.sleep(2000);

	        // co-maintanance

	    }

	    @Test(priority = 14)

	    public void ReleaseMultilineItems() throws InterruptedException,XmlException, IOException {
	        reuse= new ReusableScreenshot();
	        // click on Edit icon
	        driver.findElement(By.cssSelector("img[id*='_btadmini_table[1].thtmlb_oca.EDIT']")).click();
	        Thread.sleep(5000);
	        scroll1 = driver.findElement(By.xpath("//*[text()='Service Contract Item Details']"));
	        Actions actions = new Actions(driver);
	        actions.scrollToElement(scroll1).perform();
	        Thread.sleep(5000);
	        scroll1.click();
	        driver.findElement(By.xpath("//*[text()='Service Contract Item Details']")).click();
	        Thread.sleep(2000);
	        driver.findElement(By.xpath("//*[@id='C34_W125_V126_V127_btadmini_lcstatus']")).click();
	        Thread.sleep(2000);
	        driver.findElement(By.xpath("//*[text()='Release']")).click();
	        Thread.sleep(2000);
	        driver.findElement(By.cssSelector("a[id*='C34_W125_V126_thtmlb_button_3']")).click();
	        Thread.sleep(2000);

	        driver.findElement(By.xpath("//*[@id='C34_W125_V126_V127_btadmini_lcstatus']")).click();
	        Thread.sleep(2000);
	        driver.findElement(By.xpath("//*[text()='Release']")).click();
	        Thread.sleep(2000);
	        driver.findElement(By.cssSelector("a[id*='C34_W125_V126_thtmlb_button_3']")).click();
	        Thread.sleep(2000);
	        driver.findElement(By.xpath("//*[@id='C34_W125_V126_V127_btadmini_lcstatus']")).click();
	        Thread.sleep(2000);
	        driver.findElement(By.xpath("//*[text()='Release']")).click();
	        Thread.sleep(2000);
	        driver.findElement(By.cssSelector("a[id*='C34_W125_V126_thtmlb_button_1']")).click();
	        Thread.sleep(4000);
	        driver.findElement(By.xpath("//*[@id='C29_W88_V91_thtmlb_button_1']")).click();
	        Thread.sleep(2000);
	        WebElement text1 = driver.findElement(By.xpath("//*[@id='CRMMessageLine1']/span[3]"));
	        
	        if (text1.isDisplayed()) {
	            String textis = text1.getText();
	            textnum = textis.replaceAll("[^0-9]", "");
	            System.out.println("text is generated " + textnum);
	            TestNGResults.put("11", new Object[] { 1d, "Verifying transcation ID", "Transcation ID got created", "Pass" });
	            reuse.captureScreenshot(ScreenshotNames[array_increment++]="Transaction", driver);

	        } else {
	            System.out.println("text is not generated");
	            TestNGResults.put("11", new Object[] { 1d, "Verifying transcation ID", "Transcation ID got created", "Fail" });
	        }

	        Thread.sleep(2000);
	        driver.findElement(By.xpath("//*[@id='C4_W16_V17_SRV-CONTR']")).click();
	        Thread.sleep(2000);
	        driver.findElement(By.xpath("//*[@id='C24_W69_V70_SRV-SCO-SR']")).click();
	        Thread.sleep(3000);
	        WebElement ele1 = driver.findElement(By.xpath("//*[@id='C26_W76_V77_V78_btqsrvcon_parameters[1].VALUE1']"));
	        ele1.sendKeys(textnum.toString());

	        ele1.sendKeys(Keys.ENTER);
	        Thread.sleep(130000);
	        scroll1 = driver.findElement(By.cssSelector("a[id*='btqrsrvcon_table[1].object_id']"));
	        Actions action = new Actions(driver);
	        action.scrollToElement(scroll1).perform();
	        Thread.sleep(5000);
	        scroll1.click();
	        Thread.sleep(5000);
	        WebElement scroll2 = driver.findElement(By.xpath("//*[@id='C29_W88_V91_0014_expand_link']"));
	        Actions transaction = new Actions(driver);
	        transaction.scrollToElement(scroll2).perform();
	        Thread.sleep(3000);
	        scroll2.click();
	        Thread.sleep(3000);
	        WebElement debitmemo1 = driver.findElement(By.cssSelector("span[id*='btdocflow_table[1].process_type_des']"));
	        if (debitmemo1.isDisplayed()) {
	            System.out.println("debit memo for line item 10 is created");
	            TestNGResults.put("12", new Object[] { 1d, "Checking For Debit Memp", "Debit Memo 1 Created", "Pass" });
	        } else {
	            System.out.println("debit memo for line item 10 is not created");
	            TestNGResults.put("12", new Object[] { 1d, "Checking For Debit Memp", "Debit Memo 1 Created", "Fail" });
	        }
	        WebElement debitmemo2 = driver.findElement(By.cssSelector("span[id*='btdocflow_table[2].process_type_des']"));
	        if (debitmemo2.isDisplayed()) {
	            System.out.println("debit memo for line item 20 is created");
	            TestNGResults.put("13", new Object[] { 1d, "Checking For Debit Memp", "Debit Memo 2 Created", "Pass" });
	        } else {
	            System.out.println("debit memo for line item 20 is not created");
	            TestNGResults.put("13", new Object[] { 1d, "Checking For Debit Memp", "Debit Memo 2 Created", "fail" });

	        }
	        WebElement debitmemo3 = driver.findElement(By.cssSelector("span[id*='btdocflow_table[3].process_type_des']"));
	        if (debitmemo3.isDisplayed()) {
	            System.out.println("debit memo for line item 30 is created");
	            TestNGResults.put("14", new Object[] { 1d, "Checking For Debit Memp", "Debit Memo 3 Created", "Pass" });
	        } else {
	            System.out.println("debit memo for line item 30 is not created");
	            TestNGResults.put("14", new Object[] { 1d, "Checking For Debit Memo", "Debit Memo 3 Created", "Fail" });
	        }
	        WebElement IO1 = driver.findElement(By.cssSelector("span[id*='btdocflow_table[4].process_type_des']"));
	        if (IO1.isDisplayed()) {
	            System.out.println("IO for line item 10 is created");
	            TestNGResults.put("15", new Object[] { 1d, "Checking For IO", "IO 1 Created", "Pass" });
	        } else {
	            System.out.println("IO for line item 10 is not created");
	            TestNGResults.put("15", new Object[] { 1d, "Checking For IO", "IO 1 Created", "Fail" });
	        }
	        WebElement IO2 = driver.findElement(By.cssSelector("span[id*='btdocflow_table[5].process_type_des']"));
	        if (IO2.isDisplayed()) {
	            System.out.println("IO for line item 20 is created");
	            TestNGResults.put("16", new Object[] { 1d, "Checking For IO", "IO 2 Created", "Pass" });
	        } else {
	            System.out.println("IO for line item 20 is not created");
	            TestNGResults.put("16", new Object[] { 1d, "Checking For IO", "IO 2 Created", "Fail" });
	        }
	        WebElement IO3 = driver.findElement(By.cssSelector("span[id*='btdocflow_table[6].process_type_des']"));
	        if (IO3.isDisplayed()) {
	            System.out.println("IO for line item 30 is created");
	            TestNGResults.put("17", new Object[] { 1d, "Checking For IO", "IO 3 Created", "Pass" });
	        } else {
	            System.out.println("IO for line item 10 is not created");
	            TestNGResults.put("17", new Object[] { 1d, "Checking For IO", "IO 3 Created", "Fail" });
	        }
	        WebElement project1 = driver.findElement(By.cssSelector("span[id*='btdocflow_table[7].process_type_des']"));
	        if (project1.isDisplayed()) {
	            System.out.println("project1 is created");
	            TestNGResults.put("18", new Object[] { 1d, "Checking For Project", "Project 1 Created", "Pass" });
	            reuse.captureScreenshot(ScreenshotNames[array_increment++]="ReleaseContractLine", driver);
	        } else {
	            System.out.println("projecct1 is not created");
	            TestNGResults.put("18", new Object[] { 1d, "Checking For Project", "Project 1 Created", "Fail" });

	        }

	        Thread.sleep(3000);
	        WebElement MaintananceCase = driver.findElement(By.cssSelector("a[id*='_btdocflow_table[7].id_description']"));
	        MaintananceNO = MaintananceCase.getText();

	        WebElement project2 = driver.findElement(By.cssSelector("span[id*='btdocflow_table[8].process_type_des']"));
	        if (project2.isDisplayed()) {
	            System.out.println("project2 is created");
	            TestNGResults.put("19", new Object[] { 1d, "Checking For Project", "Project 2 Created", "Pass" });
	        } else {
	            System.out.println("project2 is not created");
	            TestNGResults.put("19", new Object[] { 1d, "Checking For Project", "Project 1 Created", "Fail" });
	        }
	        Thread.sleep(3000);
	        WebElement ServiceCase = driver.findElement(By.cssSelector("a[id*='_btdocflow_table[8].id_description']"));
	        ServiceNo = ServiceCase.getText();
	        Thread.sleep(3000);
	        driver.close();
	        Set<String> handle = driver.getWindowHandles(); // Switch to Parent window
	        Iterator<String> it = handle.iterator();
	        String Parentwindow = it.next();

	        driver.switchTo().window(Parentwindow);

	    }

	    @Test(priority = 15)
	    public void AgsProjectOrderMaintanance() throws InterruptedException,XmlException, IOException {
	        reuse= new ReusableScreenshot();
	        Thread.sleep(5000);
	        SideArrow = driver.findElement(By.xpath("//*[@id='tlnOverflowBtn'] "));
	        SideArrow.click();
	        Thread.sleep(5000);
	        SalesClk = driver.findElement(By.xpath("//*[@id='topTabMenuItem33']/td[2]/div"));
	        SalesClk.click();

	        Thread.sleep(10000);
	        agssales = driver.findElement(By.xpath("//*[@id='subTabIndex1']/div[1]"));
	        agssales.click();
	        Thread.sleep(5000);
	        Crm = driver.findElement(By.xpath("//*[text()='SAP CRM']"));
	        Crm.click();
	        Set<String> handle = driver.getWindowHandles(); // Switch to new chrome window
	        Iterator<String> it = handle.iterator();
	        String Parentwindow = it.next();
	        String Childwindow = it.next();

	        driver.switchTo().window(Childwindow);

	        Thread.sleep(5000);
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='contentAreaFrame']"))); // handling Frames
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='isolatedWorkArea']"))); // handling Frames

	        driver.findElement(By.xpath("//*[@id='ZAPG']")).click();
	        Thread.sleep(4000);
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='CRMApplicationFrame']"))); // handling Frames
	        driver.switchTo().frame(driver.findElement(By.xpath("//frame[@id ='WorkAreaFrame1']"))); // handling Frames
	        Thread.sleep(10000);
	        driver.findElement(By.cssSelector("a[id*='ZCE-ORD']")).click(); // service contract click
	        Thread.sleep(3000);
	        driver.findElement(By.xpath("//*[text()=' Search Service Orders / WDR']")).click();
	        Thread.sleep(4000);
	        WebElement ServiceOrderWDRID = driver
	                .findElement(By.cssSelector("input[id*='btqsrvord_parameters[1].VALUE1']"));
	        ServiceOrderWDRID.sendKeys("82009237");
	        ServiceOrderWDRID.sendKeys(Keys.ENTER);
	        Thread.sleep(5000);
	        driver.findElement(By.cssSelector("a[id*='_btqrsrvord_table[1].object_id']")).click();
	        Thread.sleep(10000);
	        WebElement Item2 = driver.findElement(By.xpath("//*[text()='Items']"));
	        Actions act = new Actions(driver);
	        act.scrollToElement(Item2).perform();
	        Thread.sleep(3000);
	        reuse.captureScreenshot(ScreenshotNames[array_increment++]="ProjectOrderMaintanance", driver);

	        WebElement ProductIdv = driver.findElement(By.cssSelector("a[id*='btadmini_table[1].ordered_prod']"));
	        String ProductIDM = ProductIdv.getText();
	        String ActualProductIDM = "A8192";
	        softAssert.assertEquals(ProductIDM, ActualProductIDM);
	        TestNGResults.put("19", new Object[] { 1d, "Verifying Service Case", "Service Case Displaying for ProjectOrderMaintanance ", "Pass" });


	       

	        

	        driver.close();
	        driver.switchTo().window(Parentwindow);
	    }

	    @Test(priority = 16)
	    public void AgsProjectOrderService() throws InterruptedException,XmlException, IOException {
	        reuse= new ReusableScreenshot();

	        Thread.sleep(10000);
	        Crm = driver.findElement(By.xpath("//*[text()='SAP CRM']"));
	        Crm.click();
	        Set<String> handle = driver.getWindowHandles(); // Switch to new chrome window
	        Iterator<String> it = handle.iterator();
	        String Parentwindow = it.next();
	        String Childwindow = it.next();
	        driver.switchTo().window(Childwindow);

	        Thread.sleep(5000);
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='contentAreaFrame']"))); // handling Frames
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='isolatedWorkArea']"))); // handling Frames

	        driver.findElement(By.xpath("//*[@id='ZAPG']")).click();
	        Thread.sleep(4000);
	        driver.switchTo().frame(driver.findElement(By.xpath("//iframe[@id='CRMApplicationFrame']"))); // handling Frames
	        driver.switchTo().frame(driver.findElement(By.xpath("//frame[@id ='WorkAreaFrame1']"))); // handling Frames
	        Thread.sleep(10000);
	        driver.findElement(By.cssSelector("a[id*='ZCE-ORD']")).click(); // service contract click
	        Thread.sleep(3000);
	        driver.findElement(By.xpath("//*[text()=' Search Service Orders / WDR']")).click();
	        Thread.sleep(2000);
	        WebElement ServiceOrderWDRID1 = driver
	                .findElement(By.cssSelector("input[id*='btqsrvord_parameters[1].VALUE1']"));
	        ServiceOrderWDRID1.sendKeys("82009237");
	        ServiceOrderWDRID1.sendKeys(Keys.ENTER);
	        Thread.sleep(5000);
	        driver.findElement(By.cssSelector("a[id*='_btqrsrvord_table[1].object_id']")).click();
	        Thread.sleep(5000);
	        WebElement Item2 = driver.findElement(By.xpath("//*[text()='Items']"));
	        Actions act1 = new Actions(driver);
	        act1.scrollToElement(Item2).perform();
	        Thread.sleep(3000);

	        WebElement ProductIdv = driver.findElement(By.cssSelector("a[id*='btadmini_table[1].ordered_prod']"));
	        String ProductIDM = ProductIdv.getText();
	        String ActualProductIDM = "A8192";
	        softAssert.assertEquals(ActualProductIDM, ProductIDM);
	        TestNGResults.put("21", new Object[] { 1d, "Verifying Service Case", "Service Case Displaying for ProjectOrderService ", "Pass" });
	        reuse.captureScreenshot(ScreenshotNames[array_increment++]="ProjectOrderService", driver);

	        driver.close();
	        driver.switchTo().window(Parentwindow);

	    }

	    @Test(priority = 17)
	    public void UI5ProjectOrderService() throws InterruptedException,XmlException, IOException {
	        Thread.sleep(5000);
	        SideArrow = driver.findElement(By.cssSelector("div[id*='OverflowBtn']"));
	        SideArrow.click();
	        Thread.sleep(10000);
	        UI5 = driver.findElement(By.xpath("//*[@id='topTabMenuItem35']/td[2]/div")); // Click on UI5
	        UI5.click();
	        Thread.sleep(10000);
	        Set<String> handle = driver.getWindowHandles(); // Switch to new chrome window
	        Iterator<String> it = handle.iterator();
	        String Parentwindow = it.next();
	        String Childwindow = it.next();
	        System.out.println(Parentwindow);
	        System.out.println(Childwindow);
	        driver.switchTo().window(Childwindow);
	        String r = driver.getTitle();
	        System.out.println(r);
	        Thread.sleep(5000);
	        reuse= new ReusableScreenshot();

	        ProjectMaint = driver.findElement(By.xpath("//*[@id='__tile18']"));
	        Actions action1 = new Actions(driver); // Scroll to ProjectMaintainceIcon
	        action1.scrollToElement(ProjectMaint).perform();
	        Thread.sleep(2000);
	        ProjectMaint.click();
	        Thread.sleep(10000);
	        driver.findElement(By.xpath("//*[text()='Advanced Search']")).click();
	        Thread.sleep(5000);

	        driver.findElement(By.xpath("//*[@placeholder='Enter Project Number']")).sendKeys(MaintananceNO);
	        Thread.sleep(2000);
	        driver.findElement(By.xpath("//*[text()='Search']")).click();
	        Thread.sleep(2000);

	        Thread.sleep(2000);
	        reuse.captureScreenshot(ScreenshotNames[array_increment++]="ProjectOrderMTUnderUI", driver);
	        // Milstone
	        try {
	            WebElement Milstone = driver.findElement(By.xpath("//*[@id='__filter4-text']"));
	            System.out.println("Milestone data  found for Project Order-MT" + " Pass");

	        } catch (NoSuchElementException exception) {
	            System.out.println("Milestone data not found for Project Order-MT" + " Fail");

	        }
	        // Link Pre PO
	        try {
	            WebElement LinkPrePO = driver.findElement(By.xpath("//*[@id='__filter4-text']"));
	            System.out.println("Link Pre PO  found for Project Order-MT" + " Pass");

	        } catch (NoSuchElementException exception) {
	            System.out.println("Link Pre PO data not found for Project Order-MT" + " Fail");

	        }

	        File desscreenshot = new File(
	                "C:\\Users\\X0143782\\OneDrive - Applied Materials\\Documents\\ScreenShot\\ss1.png");
	        // Copy the file to a location and use try catch block to handle exception

	        try {
	            File screenshot = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

	            FileUtils.copyFile(screenshot, desscreenshot);
	        } catch (IOException e) {

	            System.out.println(e.getMessage());

	        }
	        driver.close();
	        driver.switchTo().window(Parentwindow);

	    }

	    @Test(priority = 18)
	    public void UI5ProjectOrderMT() throws InterruptedException, IOException, XmlException {
	        reuse= new ReusableScreenshot();

	        Thread.sleep(5000);
	        SideArrow = driver.findElement(By.cssSelector("div[id*='OverflowBtn']"));
	        SideArrow.click();
	        Thread.sleep(10000);
	        UI5 = driver.findElement(By.xpath("//*[@id='topTabMenuItem35']/td[2]/div"));
	        UI5.click();
	        Thread.sleep(5000);
	        Set<String> handle = driver.getWindowHandles(); // Switch to new chrome window
	        Iterator<String> it = handle.iterator();
	        String Parentwindow = it.next();
	        String Childwindow = it.next();
	        System.out.println(Parentwindow);
	        System.out.println(Childwindow);

	        driver.switchTo().window(Childwindow);

	        Thread.sleep(5000);

	        WebElement Service = driver.findElement(By.xpath("//*[@id='__tile18']"));
	        Actions action12 = new Actions(driver);
	        action12.scrollToElement(Service).perform();
	        Thread.sleep(2000);
	        Service.click();
	        Thread.sleep(10000);
	        driver.findElement(By.xpath("//*[text()='Advanced Search']")).click(); // Click on advance search
	        Thread.sleep(5000);

	        driver.findElement(By.xpath("//*[@placeholder='Enter Project Number']")).sendKeys(ServiceNo);
	        Thread.sleep(2000);
	        driver.findElement(By.xpath("//*[text()='Search']")).click();
	        Thread.sleep(2000);

	        Thread.sleep(2000);
	        reuse.captureScreenshot(ScreenshotNames[array_increment++]="ProjectOrderServiceUnderUI5", driver);
	        // Milstone
	        try {
	            Milstone1 = driver.findElement(By.xpath("//*[@id='__filter4-text']"));
	            System.out.println("Milestone data  found for Project Order-Service" + " Fail");

	        } catch (NoSuchElementException exception) {
	            System.out.println("Milestone data not found for Project OrderService" + "Pass");

	        }
	        // Link Pre PO
	        try {
	            WebElement LinkPrePO1 = driver.findElement(By.xpath("//*[@id='__filter4-text']"));
	            System.out.println("Link Pre PO  found for Project Order-Service" + " Fail");

	        } catch (NoSuchElementException exception) {
	            System.out.println("Link Pre PO data not found for Project Order-Service" + " Pass");

	        }
	        
	        // Click on Milestones tile
	       /*  driver.findElement(By.xpath("//div[@id='__filter4-text']")).click();
	        WebElement mile1 = driver.findElement(By.xpath("//div[@id='__filter4-text']"));
	        if (mile1.isDisplayed()) {
	            System.out.println("milestone tile is Displayed");
	        } else {
	            System.out.println("milestone tile is not displayed");
	        }
	        // Verify original billing date and current billing date
	        curbill = driver.findElement(By.cssSelector("div[id*='__vbox85-__clone2']"));
	        if (curbill.isDisplayed()) {
	            String curbill1 = curbill.getText();
	            System.out.println("Current billing date is displayed " + curbill1);
	        } else {
	            System.out.println("Current billing date is not displayed");
	        }
	        orgbill = driver.findElement(By.cssSelector("div[id*='__vbox86-__clone2']"));
	        if (orgbill.isDisplayed()) {
	            String orgbill1 = curbill.getText();
	            System.out.println("original billing date is displayed " + orgbill1);
	        } else {
	            System.out.println("original billing date is not displayed");
	        }
	        // click on pencil icon
	        Thread.sleep(2000);
	        driver.findElement(By.cssSelector("span[id*='__button36-__clone2-inner']")).click();
	        verifyorignalbil = driver.findElement(By.cssSelector("div[id*='__vbox36-__clone3']"));
	        if (verifyorignalbil.isDisplayed()) {
	            String verifycurbil = curbill.getText();
	            System.out.println("Current billing date is displayed " + verifycurbil);
	        } else {
	            System.out.println("Current billing date is not displayed");
	        }
	        reuse.saveScreenshotsToWordDocument("Regression_Results",ScreenshotNames);

	        driver.quit();*/

	    }
	    @BeforeClass(alwaysRun = true)
		public void suiteSetUp() {

			// create a new work book
			workbook = new XSSFWorkbook();
			// create a new work sheet
			sheet = workbook.createSheet("TestNG Result Summary");
			TestNGResults = new LinkedHashMap<String, Object[]>();
			// add test result excel file column header
			// write the header in the first row
			TestNGResults.put("1", new Object[] { "Test Step No.", "Action", "Expected Output", "Actual Output" });

	        // Get current working directory and load the data file
				//workingDir = "C:\\Users\\X0143782\\OneDrive - Applied Materials\\Documents\\ScreenShot";

			

		}

		@AfterClass
		public void suiteTearDown() throws IOException, XmlException {
	        reuse.saveScreenshotsToWordDocument("Regression_Results",ScreenshotNames);
			// write excel file and file name is SaveTestNGResultToExcel.xls
			Set<String> keyset = TestNGResults.keySet();
			int rownum = 0;
			for (String key : keyset) {
				XSSFRow row = sheet.createRow(rownum++);
				Object[] objArr = TestNGResults.get(key);
				int cellnum = 0;
				for (Object obj : objArr) {
					XSSFCell cell = row.createCell(cellnum++);
					if (obj instanceof Date)
						cell.setCellValue((Date) obj);
					else if (obj instanceof Boolean)
						cell.setCellValue((Boolean) obj);
					else if (obj instanceof String)
						cell.setCellValue((String) obj);
					else if (obj instanceof Double)
						cell.setCellValue((Double) obj);
				}
			}
			try {
				FileOutputStream out = new FileOutputStream(new File("C:\\Users\\X0143782\\OneDrive - Applied Materials\\Documents\\ScreenShot\\SaveTestNGResultToExcel.xlsx"));
				workbook.write(out);
				out.close();
				System.out.println("Successfully saved Selenium WebDriver TestNG result to Excel File!!!");

			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
			// close the browser
			
		}

	}

}
