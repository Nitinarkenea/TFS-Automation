package mainjava;

import java.awt.AWTException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Reporter;





public class DashboardPage extends BaseClass {

	@FindBy(xpath = "//*[@id='details-button']")
    WebElement advance;
    @FindBy(xpath = "//*[@id='proceed-link']")
    WebElement proceedlink;
	//3rd PartFigure 
    @FindBy(xpath = "//*[@id='catDropdown']//li//a[text()='3rd Party Figures']")
    WebElement thirdparty;
    @FindBy(xpath = "//strong[text()='3rd Party Figures']")
    WebElement thridPartyFigure;

    //firstorder
    @FindBy(xpath = "//*[@id='filtered-products']//div[2]//div[1]//div//div[1]//a")
    WebElement firstorder;
    //addcart
    @FindBy(xpath = "//input[@name='add_carrt']")
    WebElement addcart;
    //checkout
    @FindBy(xpath = "//*[text()='Proceed to Checkout']")
	WebElement Checkout;
    //close popup
    @FindBy(xpath = "//*[@class='modal_close']")
    WebElement pop;
    
  //continue shopping
  	@FindBy(xpath = "//a[text()='Continue Shopping']")
  	WebElement contineshopping;
  	
  	@FindBy(xpath = "//*[@id='catDropdown']//li//a[text()='Japanese']")
  	WebElement japanese;
  	@FindBy(xpath = "//strong[text()='Japanese Transformers']")
  	WebElement japaneseTransformers;
  	@FindBy(xpath = "//*[@id='filtered-products']//div[2]//div[2]//div//div[1]//a")
  	WebElement secondorder;
  	
  	
  	@FindBy(xpath = "//*[@id='catDropdown']//li//a[text()='Sale']")
  	WebElement sale;
  	@FindBy(xpath = "//strong[text()='Sale Clearance']")
  	WebElement SaleTransformers;
  	@FindBy(xpath = "//*[@id='filtered-products']//div[2]//div[1]//div//div[1]//a")
  	WebElement lastorder;
  	
  	
  	
  	@FindBy(xpath = "//*[@id='catDropdown']//li//a[text()='Hasbro']")
  	WebElement Hasbro;
  	@FindBy(xpath = "//strong[text()='Hasbro Transformers']")
  	WebElement hasbrotransformer;
  	@FindBy(xpath = "//*[@id='filtered-products']//div[2]//div[3]//div//div[1]//a")
  	WebElement thirdOrder;
  	
  	@FindBy(xpath = "//*[@id='catDropdown']//li//a[text()='Vintage']")
  	WebElement vintage;
  	@FindBy(xpath = "//strong[text()='Vintage Transformers']")
  	WebElement vintagetransformer;
  	@FindBy(xpath = "//*[@id='filtered-products']//div[2]//div[2]//div//div[1]//a")
  	WebElement FourthOrder;
  	
  	@FindBy(xpath = "//*[@id='catDropdown']//li//a[text()='Exclusives']")
  	WebElement exclusive;
  	@FindBy(xpath = "//strong[text()='Transformers Exclusives']")
  	WebElement exclusivetransformer;
  	@FindBy(xpath = "//*[@id='filtered-products']//div[2]//div[2]//div//div[1]//a")
  	WebElement FifthOrder;
   //proceed to checkout
  	@FindBy(xpath = "//a[text()='Proceed to Checkout']")
  	WebElement checkout;
  	@FindBy(id = "username")
  	WebElement user;
  	@FindBy(id = "password")
  	WebElement password;
  	@FindBy(id = "signIn")
  	WebElement signIn;
  	
  	@FindBy(xpath = "//*[@class='col-6 shipping-billing-address']//p//a")
	WebElement address;
	@FindBy(xpath ="//*[@title=\"Add New Shipping Address\"]")
	WebElement newaddress;
	@FindBy(xpath = "//*[@id='lb_ship_first_name']")
	WebElement NDF_Name;
	@FindBy(xpath = "//*[@id='lb_ship_last_name']")
	WebElement NDL_Name;
	@FindBy(xpath = "//*[@id='lb_ship_address1']")
	WebElement NDA_Name;
	@FindBy(xpath = "//*[@id='lb_ship_address2']")
	WebElement NDA2_Name;
	@FindBy(xpath = "//*[@id='lb_ship_country']")
	WebElement NDC_Name;
	
	@FindBy(xpath = "//*[@id='lb_ship_zip']")
	WebElement NDZ_Name;
	@FindBy(xpath = "//*[@id='lb_ship_city']")
	WebElement NDCity_Name;
	@FindBy(xpath = "//*[@id='lb_ship_state']")
	WebElement NDState_Name;
	@FindBy(xpath = "//*[@id='lb_ship_phone']")
	WebElement NDPhone_Name;
	@FindBy(xpath = "//span[contains(text(),'Use this address as my billing information')]")
	WebElement billingAddress;
	@FindBy(xpath = "//*[@id='shipping_update']")
	WebElement ship;
	
	@FindBy(xpath = "//*[@id='add_to_stack']")
	WebElement ChooseOption;
	
	@FindBy(xpath = "//*[@id='checkout_submit']")
	WebElement PlaceOrder;
  
	
	//Payment method
	@FindBy(xpath = "//*[@id='third_edit']")
	WebElement debitcard;
	
	@FindBy(xpath = "//*[@id='customer_card_s_90']//span")
	WebElement addcard;
	@FindBy(xpath = "//*[@id='cardNumber']")
	WebElement cardnumber;
	@FindBy(xpath = "//*[@id='expMonthYear']")
	WebElement expdate;
	@FindBy(xpath = "//*[@id='cardCode']")
	WebElement cvvcode;
	@FindBy(xpath = "//*[@id='bill_first_name']")
	WebElement cardholderfirstname;
	@FindBy(xpath = "//*[@id='bill_last_name']")
	WebElement cardholderlastname;
	@FindBy(xpath = "//*[@id='bill_address1']")
	WebElement add1;
	@FindBy(xpath = "//*[@id='bill_city']")
	WebElement billCity;
	@FindBy(xpath = "//*[@id='bill_country']")
	WebElement billCountry;
	@FindBy(xpath = "//*[@id='bill_zip']")
	WebElement billZIP;
	@FindBy(xpath = "//*[@id='billing_form']//section//div/div//div//div[1]//div//div[9]//div[4]//div//div//div//input")
	WebElement place;
	@FindBy(xpath = "//*[@id='bill_state_2']")
	WebElement paymentstate;
  	
    public DashboardPage(WebDriver driver)
    {
    	this.driver=driver;
    	PageFactory.initElements(driver,this);
    }
    
    public void thirdPartFigures() throws InterruptedException
	{
    	WebDriverWait wait = new WebDriverWait(driver, TimeUnit.SECONDS.toMillis(5));
    	advance.click();
    	proceedlink.click();
    	
		Actions act = new Actions(driver);
		act.moveToElement(thirdparty).perform();
		   Thread.sleep(1000);
		thridPartyFigure.click();
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollBy(0,650)");
	    Thread.sleep(1000);
		List<WebElement> sorting=driver.findElements(By.xpath("//*[@id='sort-by']//optgroup//option"));
		for(WebElement sort:sorting)
		{
			if(sort.getText().equals("Price: Lowest First"))
			{
				sort.click();
			}
		}
		Thread.sleep(5000);
		wait.until(ExpectedConditions.elementToBeClickable(firstorder));
		firstorder.click();
		
		addcart.click();
		Reporter.log("First product selected from ThirdParty Figure", true);
	
		Thread.sleep(3000);
		
		
	}
    public void verifyJapanese() throws InterruptedException
	{
    	WebDriverWait wait1 = new WebDriverWait(driver, TimeUnit.SECONDS.toMillis(10));
		wait1.until(ExpectedConditions.elementToBeClickable(japanese));
		Actions act = new Actions(driver);
		act.moveToElement(japanese).perform();
		japaneseTransformers.click();
		
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		List<WebElement> sorting=driver.findElements(By.xpath("//*[@id='sort-by']//optgroup//option"));
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollBy(0,950)");
		Thread.sleep(2000);
		for(WebElement sort:sorting)
		{
			if(sort.getText().equals("Price: Lowest First"))
			{
				sort.click();
			}
		}
		
		
		Thread.sleep(5000);
		wait1.until(ExpectedConditions.elementToBeClickable(secondorder));
		secondorder.click();
		
	    addcart.click();
	    Reporter.log("Second product selected from Japanese", true);
	    Thread.sleep(3000);
				
}
    
    public void verifyHasbro() throws InterruptedException
	{
    	WebDriverWait wait1 = new WebDriverWait(driver, TimeUnit.SECONDS.toMillis(10));
		wait1.until(ExpectedConditions.elementToBeClickable(Hasbro));
		Actions act = new Actions(driver);
		act.moveToElement(Hasbro).perform();
		hasbrotransformer.click();
		
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		List<WebElement> sorting=driver.findElements(By.xpath("//*[@id='sort-by']//optgroup//option"));
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollBy(0,950)");
		Thread.sleep(1000);
		for(WebElement sort:sorting)
		{
			if(sort.getText().equals("Price: Lowest First"))
			{
				sort.click();
			}
		}
		Thread.sleep(4000);
		
		WebDriverWait wait = new WebDriverWait(driver, TimeUnit.SECONDS.toMillis(20));
		wait.until(ExpectedConditions.elementToBeClickable(thirdOrder));
		
		thirdOrder.click();
		
	    addcart.click();
	    Reporter.log("Third product selected from HasBro", true);
	    Thread.sleep(3000);
				
}
    public void verifyVintage() throws InterruptedException
	{
    	driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
    	WebDriverWait wait1 = new WebDriverWait(driver, TimeUnit.SECONDS.toMillis(10));
		wait1.until(ExpectedConditions.elementToBeClickable(vintage));
		Actions act = new Actions(driver);
		act.moveToElement(vintage).perform();
		vintagetransformer.click();
		
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		List<WebElement> sorting=driver.findElements(By.xpath("//*[@id='sort-by']//optgroup//option"));
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollBy(0,950)");
		Thread.sleep(2000);
		for(WebElement sort:sorting)
		{
			if(sort.getText().equals("Price: Lowest First"))
			{
				sort.click();
			}
		}
		Thread.sleep(5000);
		
		WebDriverWait wait = new WebDriverWait(driver, TimeUnit.SECONDS.toMillis(10));
		wait.until(ExpectedConditions.elementToBeClickable(FourthOrder));
		
		FourthOrder.click();
		
	    addcart.click();
	    Reporter.log("Fourth product selected from Vintage", true);
	    Thread.sleep(3000);
				
}
    
    public void verifyExclusives() throws InterruptedException
	{
    	WebDriverWait wait1 = new WebDriverWait(driver, TimeUnit.SECONDS.toMillis(10));
		wait1.until(ExpectedConditions.elementToBeClickable(exclusive));
		Actions act = new Actions(driver);
		act.moveToElement(exclusive).perform();
		exclusivetransformer.click();
		
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		List<WebElement> sorting=driver.findElements(By.xpath("//*[@id='sort-by']//optgroup//option"));
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollBy(0,1050)");
		Thread.sleep(1000);
		for(WebElement sort:sorting)
		{
			if(sort.getText().equals("Price: Lowest First"))
			{
				sort.click();
			}
		}
		Thread.sleep(4000);
		
		WebDriverWait wait = new WebDriverWait(driver, TimeUnit.SECONDS.toMillis(20));
		wait.until(ExpectedConditions.elementToBeClickable(FifthOrder));
		
		FifthOrder.click();
		
	    addcart.click();
	    Reporter.log("Fifth product selected from Exclusive", true);
	    Thread.sleep(3000);
				
}
    
    public void verifySale() throws InterruptedException
   	{
       	WebDriverWait wait1 = new WebDriverWait(driver, TimeUnit.SECONDS.toMillis(10));
   		wait1.until(ExpectedConditions.elementToBeClickable(sale));
   		Actions act = new Actions(driver);
   		act.moveToElement(sale).perform();
   		SaleTransformers.click();
   		
   		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
   		List<WebElement> sorting=driver.findElements(By.xpath("//*[@id='sort-by']//optgroup//option"));
   		JavascriptExecutor js = (JavascriptExecutor) driver;
   		js.executeScript("window.scrollBy(0,950)");
   		Thread.sleep(1000);
   		for(WebElement sort:sorting)
   		{
   			if(sort.getText().equals("Price: Lowest First"))
   			{
   				sort.click();
   			}
   		}
   		Thread.sleep(4000);
   		
   		WebDriverWait wait = new WebDriverWait(driver, TimeUnit.SECONDS.toMillis(20));
   		wait.until(ExpectedConditions.elementToBeClickable(lastorder));
   		
   		lastorder.click();
   		
   	    addcart.click();
   	    Reporter.log("Sixth product selected from Sale", true);
   	    Thread.sleep(3000);
   		
   
   }
    
    public void verifyCheckOut() throws IOException
    {
    	
    	checkout.click();
    	FileInputStream loader = new FileInputStream("C:\\Users\\Admin\\eclipse-workspace\\TFSiteAutomation\\config.properties");
		Properties prop = new Properties();
		prop.load(loader);
    	user.sendKeys(prop.getProperty("UserName"));
    	password.sendKeys(prop.getProperty("Password"));
    	  	signIn.click();
    	  	Reporter.log("Added all items into Cart", true);
    }
    
    
    
    public void chooseShippingOption() throws InterruptedException
	{
		
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollBy(0,200)");
		Thread.sleep(3000);
		ChooseOption.click();	
		Thread.sleep(1000);
		Reporter.log("Choose Shiping option-All items to my stack", true);
	}
    
    public void paymentmethod()
    {
    	JavascriptExecutor js = (JavascriptExecutor) driver;
    	js.executeScript("window.scrollBy(0,1250)");
    	debitcard.click();
    	//driver.findElement(By.xpath("//*[@id='customer_card_tr_347']//span[@class='radiobtn']")).click();
    	driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
    	
    	addcard.click();
    	cardnumber.sendKeys("4012888818888");
    	expdate.sendKeys("1032");
    	cvvcode.sendKeys("902");
    	cardholderfirstname.sendKeys("Nitin	");
    	cardholderlastname.sendKeys("Kotwal");
    	add1.sendKeys("Baner");
    	billCity.sendKeys("Pune");
    	Select sel=new Select(billCountry);
    	sel.selectByVisibleText("United States ");
    	
    	Select selstate= new Select(paymentstate);
    	selstate.selectByVisibleText("California");
    	
    	billZIP.sendKeys("46214");
    	place.click();
    	
    	
    	}
   
    public void placeOrder() throws InterruptedException, AWTException{
		
    	Thread.sleep(3000);
    	Actions act= new Actions(driver);
    	act.click(PlaceOrder).build().perform();;
    	Thread.sleep(1000);
    
		}
    
    
   

	public void shipmyStack() throws InterruptedException, AWTException, IOException
    {
    	driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
    
    	FileInputStream loader = new FileInputStream("C:\\Users\\Admin\\eclipse-workspace\\TFSiteAutomation\\config.properties");
		Properties prop = new Properties();
		prop.load(loader);
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		address.click();
		newaddress.click();
		NDF_Name.clear();
		NDF_Name.sendKeys(prop.getProperty("FirstName"));
		NDL_Name.clear();
		NDL_Name.sendKeys(prop.getProperty("LastName"));
		NDA_Name.clear();
		NDA_Name.sendKeys(prop.getProperty("Street"));
		NDA2_Name.clear();
		NDA2_Name.sendKeys(prop.getProperty("City"));
		Thread.sleep(500);
		Select selCountry=new Select(NDC_Name);
		Thread.sleep(1000);
		selCountry.selectByVisibleText("United Kingdom ");
		NDState_Name.click();
		NDState_Name.sendKeys("Calgary");
		NDZ_Name.clear();
		NDZ_Name.sendKeys(prop.getProperty("pincode"));
		NDCity_Name.clear();
		NDCity_Name.sendKeys(prop.getProperty("Location"));		
		NDPhone_Name.clear();
		NDPhone_Name.sendKeys(prop.getProperty("MobileNumber"));
		ship.click();	
    }
    
    public void productSelection() throws InterruptedException, IOException, AWTException
    {
    	Thread.sleep(2000);
    	List<WebElement> products= driver.findElements(By.xpath("//*[@id='itemsListHold']//div[1]//div/table//tbody//tr"));
    	int i=0;
    	
 
		for( i=3;i<products.size();i++)
		{
			Thread.sleep(2000);
			driver.findElement(By.xpath("//*[@id='itemsListHold']//div[1]//div//table//tbody//tr["+i+"]//td[1]//label")).click();
				
		}
    	
		Thread.sleep(3000);
    }
   
      	public void toPlaceorder() throws InterruptedException
    
    {		
      	Thread.sleep(1000);	
    	driver.findElement(By.xpath("//*[@id='stack_submit_items']")).click();
    	Thread.sleep(3000);
    	driver.findElement(By.xpath("//*[@id='submit_order_click']")).click();
    	
    	}
      	
    
      	
    	public void orderDetails() throws InterruptedException, IOException
    	{
    		JavascriptExecutor js = (JavascriptExecutor) driver;
    		js.executeScript("window.scrollBy(0,450)");
    		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
    		String ProductName1=driver.findElement(By.xpath("//*[@id='myOrders']//div[2]//div[1]//div//div[3]//div//div[2]//div[1]//div[2]//p[1]//a")).getText();
    		String ProductName2=driver.findElement(By.xpath("//*[@id='myOrders']/div[2]/div[1]/div/div[3]/div/div[3]/div[1]/div[2]/p[1]/a")).getText();   	
    		String OrderID=driver.findElement(By.xpath("//*[@id='myOrders']//div[2]//div[1]//div//div[1]//div[1]//span")).getText();
     		String itemsubTotal=driver.findElement(By.xpath("//*[@id='myOrders']//div[2]//div[1]//div//div[2]//div[2]//div[1]//p[2]")).getText();
     		String shipping=driver.findElement(By.xpath("//*[@id='myOrders']//div[2]//div[1]/div//div[2]//div[2]//div[2]//p[2]")).getText();
      		String grandTotal=driver.findElement(By.xpath("//*[@id='myOrders']//div[2]//div[1]//div//div[2]//div[2]//div[3]//p[2]")).getText();
      		String address=driver.findElement(By.xpath("//*[@id='myOrders']/div[2]/div[1]/div/div[2]/div[1]/p[2]")).getText();
      		 File file =    new File("C:\\Users\\Admin\\Desktop\\New Microsoft Excel Worksheet.xlsx");
 	        
 	        //Create an object of FileInputStream class to read excel file
 	        FileInputStream inputStream = new FileInputStream(file);
 	        
 	        //creating workbook instance that refers to .xls file
 	        XSSFWorkbook wb=new XSSFWorkbook(inputStream);
 	        
 	        //creating a Sheet object using the sheet Name
 	        XSSFSheet sheet=wb.getSheetAt(0);
 	        
 	        //Create a row object to retrieve row at index 3
 	       
 	        XSSFRow row2=sheet.createRow(3);
 	        XSSFRow row1=sheet.createRow(2);
 	        
 	        //create a cell object to enter value in it using cell Index
 	        
 	        row1.createCell(0).setCellValue(OrderID);
 	        row1.createCell(1).setCellValue(ProductName1);
 	        row1.createCell(2).setCellValue(itemsubTotal);
 	        row1.createCell(3).setCellValue(shipping);
 	        row1.createCell(4).setCellValue(grandTotal);
 	        row1.createCell(5).setCellValue(address);
 	        row2.createCell(1).setCellValue(ProductName2);
	        row2.createCell(2).setCellValue(itemsubTotal);
	        row2.createCell(3).setCellValue(shipping);
	        row2.createCell(4).setCellValue(grandTotal);
	        row2.createCell(5).setCellValue(address);
 	       
 	        
 	        //write the data in excel using output stream
 	        FileOutputStream outputStream = new FileOutputStream("C:\\Users\\Admin\\Desktop\\New Microsoft Excel Worksheet.xlsx");
 	        wb.write(outputStream);
 	        wb.close();
    		}
    	
    	public void orderDetails1() throws InterruptedException, IOException
    	{
    		JavascriptExecutor js = (JavascriptExecutor) driver;
    		js.executeScript("window.scrollBy(0,450)");
    		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
    		String ProductName1=driver.findElement(By.xpath("//*[@id='myOrders']//div[2]//div[1]//div//div[3]//div//div[2]//div[1]//div[2]//p[1]//a")).getText();
    		String ProductName2=driver.findElement(By.xpath("//*[@id='myOrders']/div[2]/div[1]/div/div[3]/div/div[3]/div[1]/div[2]/p[1]/a")).getText();   	
    		  	
    		String OrderID=driver.findElement(By.xpath("//*[@id='myOrders']//div[2]//div[1]//div//div[1]//div[1]//span")).getText();
     		String itemsubTotal=driver.findElement(By.xpath("//*[@id='myOrders']//div[2]//div[1]//div//div[2]//div[2]//div[1]//p[2]")).getText();
     		String shipping=driver.findElement(By.xpath("//*[@id='myOrders']//div[2]//div[1]/div//div[2]//div[2]//div[2]//p[2]")).getText();
      		String grandTotal=driver.findElement(By.xpath("//*[@id='myOrders']//div[2]//div[1]//div//div[2]//div[2]//div[3]//p[2]")).getText();
      		String address=driver.findElement(By.xpath("//*[@id='myOrders']/div[2]/div[1]/div/div[2]/div[1]/p[2]")).getText();
      		 File file =    new File("C:\\Users\\Admin\\Desktop\\New Microsoft Excel Worksheet.xlsx");
 	        
 	        //Create an object of FileInputStream class to read excel file
 	        FileInputStream inputStream = new FileInputStream(file);
 	        
 	        //creating workbook instance that refers to .xls file
 	        XSSFWorkbook wb=new XSSFWorkbook(inputStream);
 	        
 	        //creating a Sheet object using the sheet Name
 	        XSSFSheet sheet=wb.getSheetAt(0);
 	        
 	        //Create a row object to retrieve row at index 3
 	       
 	        XSSFRow row2=sheet.createRow(5);
 	        XSSFRow row1=sheet.createRow(4);
 	        
 	        //create a cell object to enter value in it using cell Index
 	        
 	        row1.createCell(0).setCellValue(OrderID);
 	        row1.createCell(1).setCellValue(ProductName1);
 	        row1.createCell(2).setCellValue(itemsubTotal);
 	        row1.createCell(3).setCellValue(shipping);
 	        row1.createCell(4).setCellValue(grandTotal);
 	        row1.createCell(5).setCellValue(address);
 	        row2.createCell(1).setCellValue(ProductName2);
	        row2.createCell(2).setCellValue(itemsubTotal);
	        row2.createCell(3).setCellValue(shipping);
	        row2.createCell(4).setCellValue(grandTotal);
	        row2.createCell(5).setCellValue(address);
 	       
 	        
 	        //write the data in excel using output stream
 	        FileOutputStream outputStream = new FileOutputStream("C:\\Users\\Admin\\Desktop\\New Microsoft Excel Worksheet.xlsx");
 	        wb.write(outputStream);
 	        wb.close();
    		}
    	public void repetation() throws InterruptedException, AWTException, IOException
      	{
    			Thread.sleep(3000);
	      		driver.findElement(By.xpath("//*[text()='Stack']")).click();
	      		Thread.sleep(2000);
	      		shipmyStack2();
	      		Thread.sleep(3000);
	      		paymentOption2();
	      		Thread.sleep(3000);
	      		toPlaceorder();
	      		Thread.sleep(3000);
	      		
	      		orderDetails1();
	      		Thread.sleep(2000);
	  	
      	}
    	
    	public void shipmyStack2() throws InterruptedException, AWTException, IOException
        {
        	        	
        	driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
           	FileInputStream loader = new FileInputStream("C:\\Users\\Admin\\eclipse-workspace\\TFSiteAutomation\\config.properties");
    		Properties prop = new Properties();
    		prop.load(loader);
    		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
    		address.click();
    		newaddress.click();
    		NDF_Name.clear();
    		NDF_Name.sendKeys(prop.getProperty("fname"));
    		NDL_Name.clear();
    		NDL_Name.sendKeys(prop.getProperty("lname"));
    		NDA_Name.clear();
    		NDA_Name.sendKeys(prop.getProperty("street"));
    		NDA2_Name.clear();
    		NDA2_Name.sendKeys(prop.getProperty("city"));
    		Select selCountry=new Select(NDC_Name);
    		selCountry.selectByVisibleText("India");
    		NDZ_Name.clear();
    		NDZ_Name.sendKeys(prop.getProperty("pinCode"));
    		NDCity_Name.clear();
    		NDCity_Name.sendKeys(prop.getProperty("location"));
    		NDState_Name.sendKeys("Karnataka");
    		NDPhone_Name.clear();
    		NDPhone_Name.sendKeys(prop.getProperty("MobNo"));
    		ship.click();
    	
        }
    	
    	
    	public void paymentOption2() throws InterruptedException
    	{
    		JavascriptExecutor js = (JavascriptExecutor) driver;
    		js.executeScript("window.scrollBy(0,1050)");
    		driver.findElement(By.xpath("//*[@class='col-6 shipping-billing-address payment']//p//a")).click();
    		Thread.sleep(2000);
    		driver.findElement(By.xpath("//input[@id='payment_type_2']")).click();
    		
    	}
    	
    	
    	public void logout()
    	{
    		WebElement myaccount=driver.findElement(By.xpath("//*[text()='My Source']"));
    		Actions act = new Actions(driver);
    		act.moveToElement(myaccount).perform();
    		driver.findElement(By.xpath("//*[@id=\"secMenu\"]/li[1]/ul/li[5]/a")).click();
    		Reporter.log("Logout Successfully",true);
    	}
        
}

