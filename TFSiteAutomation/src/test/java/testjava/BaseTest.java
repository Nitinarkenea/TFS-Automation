package testjava;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;

import io.github.bonigarcia.wdm.WebDriverManager;
import mainjava.BaseClass;
import mainjava.DashboardPage;
import mainjava.SignUpPage;


public class BaseTest extends BaseClass{

	@BeforeSuite
	public void initializingWeb() throws InterruptedException
	{
		WebDriverManager.chromedriver().setup();
		ChromeOptions options = new ChromeOptions();
		options.addArguments("--start-maximized");
		driver= new ChromeDriver(options);
		
		driver.get("https://Arkenea:Arkenea%40pk@ec2-52-43-180-42.us-west-2.compute.amazonaws.com");
		
		driver.manage().window().maximize();
		
		
	}
	
	
	@BeforeMethod
	public void allPageObject() throws IOException
	{
		
		dash= new DashboardPage(driver);
		 signup=new SignUpPage(driver);
		
		
						
	}
	
	@AfterSuite
	public void tearDowmBrowser() throws InterruptedException
	{
		Thread.sleep(2000);
		driver.close();
	}
}
