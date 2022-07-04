package mainjava;

import java.awt.AWTException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

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



public class SignUpPage extends BaseClass{
	@FindBy(xpath = "//span[text()='Account']")
	WebElement account;
	@FindBy(xpath = "//*[@id='customer_first_name']")
	WebElement firstName;
	@FindBy(xpath = "//*[@id='customer_last_name']")
	WebElement lastName;
	@FindBy(xpath = "//*[@id='customer_email']")
	WebElement email;
	@FindBy(xpath = "//*[@id='password']")
	WebElement password;
	@FindBy(xpath = "//*[@id='retype_customer_password']")
	WebElement confirmpassword;
	@FindBy(xpath = "//*[@id='customer_dob']")
	WebElement dateOfBirth;
	@FindBy(xpath = "//*[@class='ui-datepicker-month']")
	WebElement month;
	@FindBy(xpath = "//*[@class='ui-datepicker-year']")
	WebElement year;
	@FindBy(xpath = "//table//tr/td//a[text()=2]")
	WebElement date;

	@FindBy(xpath = "//*[@id='receive_sms']")

	WebElement sendTextSMS;
	@FindBy(xpath = "//button[contains(text(),'CREATE AN ACCOUNT')]")
	WebElement createNewAccount;
	@FindBy(xpath = "//input[@id='identifierId']")
	WebElement username;
	@FindBy(xpath = "//span[text()='Next']")
	WebElement next;
	@FindBy(xpath = "//input[@name='password']")
	WebElement pass;
	@FindBy(xpath = "//input[@id='receive_sms'] //ancestor::label")
	WebElement textsms;
	@FindBy(xpath = "//*[@name='country_code']")
	WebElement country;
	@FindBy(xpath = "//*[@id='mobile_number']")
	WebElement mobilenumber;

	@FindBy(xpath = "//span[text()=\"My Account\"]")
	WebElement Sign_in;
	@FindBy(xpath = "//*[@id='username']")
	WebElement user_name;
	@FindBy(xpath = "//*[@id='password']")
	WebElement SignIn_password;
	@FindBy(xpath = "//button[@id='signIn']")
	WebElement SignIn_button;
	@FindBy(xpath = "//*[@class='lab_check']")
	WebElement rememberpassword;

	@FindBy(xpath = "//button[@id='details-button']")
	WebElement advanced;
	@FindBy(id = "proceed-link")
	WebElement procced;

	@FindBy(partialLinkText = "DK-20 Upgrade Kit for Transformers Studio Series 69 Devastator Set of 8 | DNA Design")
	WebElement item;
	@FindBy(xpath = "//*[@class='btn redBtn medium add--to--btn'] //input")
	WebElement addtoCart;
	@FindBy(xpath = "//*[text()='Proceed to Checkout']")
	WebElement Checkout;

	@FindBy(xpath = "//*[@id='checkout_submit']")
	WebElement PlaceOrder;
	
	@FindBy(xpath = "//*[@class='inp_select quantity_update']")
	WebElement quantity;
	
	@FindBy(xpath = "//a[@class='modal_act ']")
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
	@FindBy(xpath = "//*[@id='lb_ship_state_2']")
	WebElement NDState_Name;
	@FindBy(xpath = "//*[@id='lb_ship_phone']")
	WebElement NDPhone_Name;
	@FindBy(xpath = "//span[contains(text(),'Use this address as my billing information')]")
	WebElement billingAddress;
	@FindBy(xpath = "//*[@id='shipping_update']")
	WebElement ship;
	
	@FindBy(xpath = "//*[@id='ship_instock']")
	WebElement ChooseOption;
	
	@FindBy(xpath = "//*[@id='third_edit']")
	WebElement debitcard;
	
	@FindBy(xpath = "//*[@id=\"customer_card_s_90\"]/span")
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
	@FindBy(xpath = "//*[@id=\"billing_form\"]/section/div/div/div/div[1]/div/div[9]/div[4]/div/div/div/input")
	WebElement place;
	
	@FindBy(xpath = "//*[@id='breadcrumb'] //li//a")
	WebElement myaccount;
	@FindBy(xpath = "//*[@class='my-source-detail'] //h3[text()='Login & Security']")
	WebElement myaccountlogin;
	@FindBy(xpath = "//*[@class='my-source-detail'] //h3[text()='Your Addresses']")
	WebElement myaccountAddresses;
	@FindBy(xpath = "//*[@class='my-source-detail'] //h3[text()='Payment Options']")
	WebElement myaccountPaymentmethods;
	@FindBy(xpath = "//*[@class='my-source-detail'] //h3[text()='Source Points']")
	WebElement myaccountSourcePoint;
	@FindBy(xpath = "//li[@class='w_submenu']//a//span")
	WebElement myaccountdetails;
	
	
	
	public SignUpPage(WebDriver driver) {
		this.driver = driver;
		PageFactory.initElements(driver, this);

	}

	public boolean signupPageTitle() {
		String title = driver.getTitle();
		if (title.equalsIgnoreCase("Collectible Transformers Toys and Action Figures | TFSource"))
			return true;
		else
			return false;
	}

//	public void createAnAccount() throws IOException, InterruptedException {
//		driver.manage().timeouts().implicitlyWait(50, TimeUnit.SECONDS);
//		account.click();
//		File file = new File("C:\Users\Admin\eclipse-workspace\TFSite\config.properties");
//		FileInputStream loader = new FileInputStream(file);
//		Properties prop = new Properties();
//		prop.load(loader);
//		
//		firstName.sendKeys(prop.getProperty("FirstName"));
//	
//		lastName.sendKeys(prop.getProperty("LastName"));
//	
//		email.sendKeys(prop.getProperty("Email"));
//		Thread.sleep(2000);
//		password.sendKeys(prop.getProperty("Password"));
//		Thread.sleep(2000);
//		confirmpassword.click();
//		confirmpassword.sendKeys(prop.getProperty("Password"));
//		Thread.sleep(2000);
//		dateOfBirth.click();
//	
//		Select selMonth = new Select(month);
//
//		selMonth.selectByVisibleText("Oct");
//
//		Select selYear = new Select(year);
//
//		selYear.selectByVisibleText("1996");
//
//		date.click();
//		Thread.sleep(1000);
//		textsms.click();
//
//		Select countrycode = new Select(country);
//
//		countrycode.selectByValue("91");
//		mobilenumber.click();
//		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
//		mobilenumber.sendKeys(prop.getProperty("MobileNo"));
//
//		createNewAccount.click();
//		
//		Reporter.log("Activation link send to EmailID",true);
//
//	}



	public void verifysignIN() throws IOException, InterruptedException {
		Thread.sleep(1000);
		FileInputStream loader = new FileInputStream("C:\\Users\\Admin\\eclipse-workspace\\TFSiteAutomation\\config.properties");
		Properties prop = new Properties();
		prop.load(loader);
		driver.get("http://Arkenea:Arkenea@pk@ec2-52-43-180-42.us-west-2.compute.amazonaws.com/");
		driver.manage().window().maximize();
		Sign_in.click();
		user_name.sendKeys(prop.getProperty("Email"));
		SignIn_password.sendKeys(prop.getProperty("Password"));
		 rememberpassword.click();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		SignIn_button.click();
		
		String titleofPage = driver.getTitle();
		System.out.println(titleofPage);
		advanced.click();
		procced.click();
		Reporter.log("SignIn successfully",true);
		driver.get("https://Arkenea:Arkenea@pk@ec2-52-43-180-42.us-west-2.compute.amazonaws.com/");
	}

	public void verfyThirdParty() {
		WebElement allCatagory = driver.findElement(By.xpath("//*[@id='catDropdown']//li//a[text()='All Categories']"));
		Actions act = new Actions(driver);
		act.moveToElement(allCatagory).perform();
		driver.findElement(By.partialLinkText("3rd Party Add ons")).click();
		
	}

	public boolean selectItemfromBestSeller() {
		// select order
		
		item.click();

		String ordertitle = driver.getTitle();
		Reporter.log(ordertitle, true);
		Reporter.log("DK-20 Upgrade Kit -Selected Item from BestSeller  ",true);
		if (ordertitle.equalsIgnoreCase(
				"DK-20 Upgrade Kit for Transformers Studio Series 69 Devastator Set of 8 | DNA Design"))
			return true;
		else
			return false;
		

	}

	public boolean verifyAddtocart() {
		// add to cart
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		addtoCart.click();
		String cartTitle = driver.getTitle();
		Reporter.log(cartTitle, false);
		Reporter.log("Item added into Cart successfully", true);
		if (cartTitle.equalsIgnoreCase("Cart Page"))
			return true;
		else
			return false;
	}

	
	public  void selectNoOfQuantity() throws InterruptedException
	{
		
		Select selQuantity=new Select(quantity);
		selQuantity.selectByVisibleText("3");
		Thread.sleep(2000);
			
		Reporter.log("Number of qualtity Selected", true);
		
	}
	public void verifyCheckout() throws InterruptedException, AWTException {

		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		Checkout.click();
		

	}
	
	
	public void selectAddress() throws InterruptedException, AWTException
	{
		
		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
		address.click();
		newaddress.click();
		NDF_Name.clear();
		NDF_Name.sendKeys("Nitin");
		NDL_Name.clear();
		NDL_Name.sendKeys("Kotwal");
		NDA_Name.clear();
		NDA_Name.sendKeys("Balaji Nagar");
		NDA2_Name.clear();
		NDA2_Name.sendKeys("Baner");
		Select selCountry=new Select(NDC_Name);
		selCountry.selectByVisibleText("India ");
		NDZ_Name.clear();
		NDZ_Name.sendKeys("32003");
		NDCity_Name.clear();
		NDCity_Name.sendKeys("Pune");
		//NDState_Name.sendKeys("Maharashtra");
		
		
		NDPhone_Name.clear();
		NDPhone_Name.sendKeys("02425325");
		billingAddress.click();
		ship.click();
		Thread.sleep(2000);
		
		
		
	}
	
	public void chooseShippingOption() throws InterruptedException
	{
		
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollBy(0,200)");
		Thread.sleep(3000);
		ChooseOption.click();
		Thread.sleep(3000);
	}

//	public void paymentmethod()
//	{
//		
//		debitcard.click();
//		//driver.findElement(By.xpath("//*[@id='customer_card_tr_347']//span[@class='radiobtn']")).click();
//		driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
//		JavascriptExecutor js = (JavascriptExecutor) driver;
//		js.executeScript("window.scrollBy(0,820)");
//		addcard.click();
//		cardnumber.sendKeys("4012888818888");
//		expdate.sendKeys("1032");
//		cvvcode.sendKeys("902");
//		cardholderfirstname.sendKeys("Nitin	");
//		cardholderlastname.sendKeys("Kotwal");
//		add1.sendKeys("Baner");
//		billCity.sendKeys("Pune");
//		Select sel=new Select(billCountry);
//		sel.selectByVisibleText("India ");
//		
//		billZIP.sendKeys("46226");
//		place.click();
//		
//		}
	
	public boolean placeOrder() throws InterruptedException, AWTException{
		
		
		
		
		Thread.sleep(2000);
		PlaceOrder.click();
		String placorder=driver.getTitle();
		Reporter.log("Order Placed succesfully", true);
		if(placorder.equals("Thank you for shopping at TF Source.com!"))
		return true;
		else
			{return false;}
		
		
		}
	
	public boolean myOrder() throws InterruptedException
	{
		
		WebElement myOrder=driver.findElement(By.xpath("//a[text()='View all your orders ']"));
		Thread.sleep(2000);
		myOrder.click();
		Thread.sleep(2000);
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollBy(0,470)");
		String ordertitle=driver.getTitle();
		Reporter.log(ordertitle,true);
		if(ordertitle.equals("TF Source - Orders History"))
		{return true;}
		else {return false;}
	
	}
	
	public boolean gotoMyAccountLogin() throws InterruptedException {
		Thread.sleep(2000);
		driver.findElement(By.xpath("//ul[@id='breadcrumb']//li[1]//a")).click();
		String titleofmyaccount=driver.getTitle();
		myaccountlogin.click();
		Thread.sleep(1000);
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollBy(0,220)");
		Reporter.log("My Account login & security details", true);
		if(titleofmyaccount.equals("TF Source - Personal Data"))
			return true;
		else
		{
			return false;
		}
	}	
	
	public boolean gotoMyAddress() throws InterruptedException {
		Thread.sleep(5000);
		driver.findElement(By.xpath("//ul[@id='breadcrumb']//li[2]//a")).click();
		myaccountAddresses.click();
		String addressdetail=driver.getTitle();
		Thread.sleep(1000);
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollBy(0,220)");
		Reporter.log("My account Addresses", true);
		if(addressdetail.equals("TF Source - Address Book Manager"))
		{return true;}
		else
		{
			return false;
			}
		
	}	
	
	public boolean gotoPayment() throws InterruptedException
	{
		Thread.sleep(5000);
		driver.findElement(By.xpath("//ul[@id='breadcrumb']//li[1]//a")).click();
		
		myaccountPaymentmethods.click();
		String spdetail=driver.getTitle();
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollBy(0,180)");
		Reporter.log("My Account Payment options", true);
		if(spdetail.equals("TF Source - Manage Card Info"))
		{return true;}
		else
		{
			return false;
			}
		
	}
	
	public boolean gotoSourcePoints() throws InterruptedException
	{
		Thread.sleep(5000);
		driver.findElement(By.xpath("//ul[@id='breadcrumb']//li[1]//a")).click();
		
		myaccountSourcePoint.click();
		String sp=driver.getTitle();
		
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollBy(0,210)");
		Reporter.log("Awarded No. Of  Source Point ", true);
		if(sp.equals("TF Source - Source Points Summary"))
		{return true;}
		else
		{
			return false;
			}
	}
	
	public boolean gotoWishList() throws InterruptedException
	{
		Thread.sleep(5000);
		WebElement myaccount=driver.findElement(By.xpath("//*[text()='My Source']"));
		Actions act = new Actions(driver);
		act.moveToElement(myaccount).perform();
		driver.findElement(By.xpath("//li[@class='w_submenu']//li[4]//a[text()='My Wishlist']")).click();
		String wishlist= driver.getTitle();
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("window.scrollBy(0,200)");
		Reporter.log("My Wishlist", true);
				if(wishlist.equals("TF Source - My Wishlist"))
			return true;
		else
			return false;
		
		
	}
	
	public boolean logout() throws InterruptedException
   {
		Thread.sleep(5000);
		WebElement myaccount=driver.findElement(By.xpath("//*[text()='My Source']"));
		Actions act = new Actions(driver);
		act.moveToElement(myaccount).perform();
		driver.findElement(By.xpath("//*[@id=\"secMenu\"]/li[1]/ul/li[5]/a")).click();
		Reporter.log("Logout Successfully",true);
		String message=driver.getTitle();
		if(message.equalsIgnoreCase("Logout Successfully - TF Source"))
		return true;
		else
			return false;
	}
	

}
