package commonFunctions;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.Properties;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.Reporter;

public class FunctionLibrary {
public static WebDriver driver;
public static Properties conpro;
//method for launch ing browser
public static WebDriver startBrowser()throws Throwable
{
	conpro = new Properties();
	conpro.load(new FileInputStream("./PropertyFiles\\Environment.properties"));
	if(conpro.getProperty("Browser").equalsIgnoreCase("chrome"))
	{
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}
	else if(conpro.getProperty("Browser").equalsIgnoreCase("firefox"))
	{
		driver = new FirefoxDriver();
	}
	else
	{
		Reporter.log("Browser value is Not Matching",true);
	}
	return driver;
}
//method for launch url
public static void openUrl()
{
	driver.get(conpro.getProperty("Url"));
}
//method for wait for any webelement 
public static void waitForElement(String LocatorType,String LocatorValue,String TestData)
{
WebDriverWait mywait = new WebDriverWait(driver, Duration.ofSeconds(Integer.parseInt(TestData)));
if(LocatorType.equalsIgnoreCase("xpath"))
{
	mywait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(LocatorValue)));
}
if(LocatorType.equalsIgnoreCase("id"))
{
	mywait.until(ExpectedConditions.visibilityOfElementLocated(By.id(LocatorValue)));
}
if(LocatorType.equalsIgnoreCase("name"))
{
	mywait.until(ExpectedConditions.visibilityOfElementLocated(By.name(LocatorValue)));
}
}
//method for any textbox
public static void typeAction(String LocatorType,String LocatorValue,String TestData)
{
	if(LocatorType.equalsIgnoreCase("xpath"))
	{
		driver.findElement(By.xpath(LocatorValue)).clear();
		driver.findElement(By.xpath(LocatorValue)).sendKeys(TestData);
	}
	if(LocatorType.equalsIgnoreCase("id"))
	{
		driver.findElement(By.id(LocatorValue)).clear();
		driver.findElement(By.id(LocatorValue)).sendKeys(TestData);
	}
	if(LocatorType.equalsIgnoreCase("name"))
	{
		driver.findElement(By.name(LocatorValue)).clear();
		driver.findElement(By.name(LocatorValue)).sendKeys(TestData);
	}
	
}
//method for click elements like buttons,images,links,checkboxes and radio button
public static void clickAction(String LocatorType,String LocatorValue)
{
	if(LocatorType.equalsIgnoreCase("xpath"))
	{
		driver.findElement(By.xpath(LocatorValue)).click();
	}
	if(LocatorType.equalsIgnoreCase("name"))
	{
		driver.findElement(By.name(LocatorValue)).click();
	}
	if(LocatorType.equalsIgnoreCase("id"))
	{
		driver.findElement(By.id(LocatorValue)).sendKeys(Keys.ENTER);
	}
}
//method for validating title
public static void validateTitle(String Expected_Title)
{
	String Actual_Title = driver.getTitle();
	try {
	Assert.assertEquals(Actual_Title, Expected_Title, "Title is Not matching");
	}catch(Throwable t)
	{
		System.out.println(t.getMessage());
	}
}
//method for closing browser
public static void closeBrowser()
{
	driver.quit();
}
//method for gneratedate
public static String generateDate()
{
	Date date = new Date();
	DateFormat df = new SimpleDateFormat("YYYY_MM_dd hh_mm_ss");
	return df.format(date);
}
//method for selecting items in listboxes
public static void dropDownAction(String LocatorType,String Locatorvalue,String TestData)
{
	if(LocatorType.equalsIgnoreCase("name"))
	{
		int value = Integer.parseInt(TestData);
		Select element = new Select(driver.findElement(By.name(Locatorvalue)));
		element.selectByIndex(value);
		
	}
	if(LocatorType.equalsIgnoreCase("xpath"))
	{
		int value = Integer.parseInt(TestData);
		Select element = new Select(driver.findElement(By.xpath(Locatorvalue)));
		element.selectByIndex(value);
	}
	if(LocatorType.equalsIgnoreCase("id"))
	{
		int value = Integer.parseInt(TestData);
		Select element = new Select(driver.findElement(By.id(Locatorvalue)));
		element.selectByIndex(value);
	}
}
//method for capture stock number into note pad
public static void captureStock(String LocatorType,String LocatorValue)throws Throwable
{
	String stockNum="";
	if(LocatorType.equalsIgnoreCase("name"))
	{
		stockNum =driver.findElement(By.name(LocatorValue)).getAttribute("value");
		
	}
	if(LocatorType.equalsIgnoreCase("xpath"))
	{
		stockNum =driver.findElement(By.xpath(LocatorValue)).getAttribute("value");
		
	}
	if(LocatorType.equalsIgnoreCase("id"))
	{
		stockNum =driver.findElement(By.id(LocatorValue)).getAttribute("value");
		
	}
	//create new notepad into capturedata folder
	FileWriter fw = new FileWriter("./CaptureData/stockNumber.txt");
	BufferedWriter bw =new BufferedWriter(fw);
	bw.write(stockNum);
	bw.flush();
	bw.close();
	
	
}
//verify stock number from stock table
public static void stockTable()throws Throwable
{
	//read stock number from above notepad
	FileReader fr = new FileReader("./CaptureData/stockNumber.txt");
	BufferedReader br = new BufferedReader(fr);
	//read stock number from notepad
	String Exp_Data =br.readLine();
	if(!driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).isDisplayed())
		driver.findElement(By.xpath(conpro.getProperty("search-panel"))).click();
	driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).clear();
	//enter stock number into textbox
	driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).sendKeys(Exp_Data);
	//click search button
	driver.findElement(By.xpath(conpro.getProperty("search-button"))).click();
	Thread.sleep(4000);
	String Act_Data = driver.findElement(By.xpath("//table[@class='table ewTable']/tbody/tr[1]/td[8]/div/span/span")).getText();
	Reporter.log(Exp_Data+"     "+Act_Data,true);
	try {
	Assert.assertEquals(Act_Data, Exp_Data, "Stock Number Not Matching");
	}catch(Throwable t)
	{
		System.out.println(t.getMessage());
	}
}
//method capture supplier number into note pad
public static void captureSup(String LocatorName,String LocatorValue)throws Throwable
{
	String supplierNum="";
	if(LocatorName.equalsIgnoreCase("xpath"))
	{
		supplierNum = driver.findElement(By.xpath(LocatorValue)).getAttribute("value");
	}
	if(LocatorName.equalsIgnoreCase("name"))
	{
		supplierNum = driver.findElement(By.name(LocatorValue)).getAttribute("value");
	}
	if(LocatorName.equalsIgnoreCase("id"))
	{
		supplierNum = driver.findElement(By.id(LocatorValue)).getAttribute("value");
	}
	//create note pad and write supplier number
	FileWriter fw = new FileWriter("./CaptureData/suppliernumber.txt");
	BufferedWriter bw = new BufferedWriter(fw);
	bw.write(supplierNum);
	bw.flush();
	bw.close();
}
//method for verifying supplier number in table
public static void supplierTable()throws Throwable
{
	//read supplier number from notepad
	FileReader fr = new FileReader("./CaptureData/suppliernumber.txt");
	BufferedReader br = new BufferedReader(fr);
	String Exp_Data = br.readLine();
	if(!driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).isDisplayed())
		//click search panel if not displayed
		driver.findElement(By.xpath(conpro.getProperty("search-panel"))).click();
	driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).clear();
	driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).sendKeys(Exp_Data);
	driver.findElement(By.xpath(conpro.getProperty("search-button"))).click();
	Thread.sleep(3000);
	String Act_Data = driver.findElement(By.xpath("//table[@class='table ewTable']/tbody/tr[1]/td[6]/div/span/span")).getText();
	Reporter.log(Exp_Data+"      "+Act_Data,true);
	try {
		Assert.assertEquals(Act_Data, Exp_Data, "Supplier Number is Not Matching");
	} catch (AssertionError a) {
		System.out.println(a.getMessage());
	}
	
	
}
//method capture customer number into note pad
public static void captureCus(String LocatorName,String LocatorValue)throws Throwable
{
	String CustomerNum="";
	if(LocatorName.equalsIgnoreCase("xpath"))
	{
		CustomerNum = driver.findElement(By.xpath(LocatorValue)).getAttribute("value");
	}
	if(LocatorName.equalsIgnoreCase("name"))
	{
		CustomerNum = driver.findElement(By.name(LocatorValue)).getAttribute("value");
	}
	if(LocatorName.equalsIgnoreCase("id"))
	{
		CustomerNum = driver.findElement(By.id(LocatorValue)).getAttribute("value");
	}
	//create note pad and write supplier number
	FileWriter fw = new FileWriter("./CaptureData/customernumber.txt");
	BufferedWriter bw = new BufferedWriter(fw);
	bw.write(CustomerNum);
	bw.flush();
	bw.close();
}
//method for verifying supplier number in table
public static void customerTable()throws Throwable
{
	//read supplier number from notepad
	FileReader fr = new FileReader("./CaptureData/customernumber.txt");
	BufferedReader br = new BufferedReader(fr);
	String Exp_Data = br.readLine();
	if(!driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).isDisplayed())
		//click search panel if not displayed
		driver.findElement(By.xpath(conpro.getProperty("search-panel"))).click();
	driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).clear();
	driver.findElement(By.xpath(conpro.getProperty("search-textbox"))).sendKeys(Exp_Data);
	driver.findElement(By.xpath(conpro.getProperty("search-button"))).click();
	Thread.sleep(3000);
	String Act_Data = driver.findElement(By.xpath("//table[@class='table ewTable']/tbody/tr[1]/td[5]/div/span/span")).getText();
	Reporter.log(Exp_Data+"      "+Act_Data,true);
	try {
		Assert.assertEquals(Act_Data, Exp_Data, "Customer Number is Not Matching");
	} catch (AssertionError a) {
		System.out.println(a.getMessage());
	}
	
	
}
}










