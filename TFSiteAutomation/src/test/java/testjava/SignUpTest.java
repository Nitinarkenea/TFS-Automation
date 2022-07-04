package testjava;

import org.testng.annotations.Test;
import org.testng.AssertJUnit;
import org.testng.annotations.Test;
import org.testng.AssertJUnit;
import org.testng.annotations.Test;
import org.testng.AssertJUnit;
import org.testng.annotations.Test;
import org.testng.AssertJUnit;
import org.testng.annotations.Test;
import org.testng.AssertJUnit;
import java.awt.AWTException;
import java.io.IOException;

import org.testng.Assert;
import org.testng.AssertJUnit;
import org.testng.annotations.Test;



public class SignUpTest extends BaseTest{

	@Test(priority=0)
	public void verifyTitleOfSignUpPage()
	{
		boolean title=signup.signupPageTitle();
		//AssertJUnit.assertEquals(title, true);
	
	}

	@Test(priority=1)
	public void verifysignInPage() throws IOException, InterruptedException
	{
		signup.verifysignIN();
	}

	@Test(priority=2)
	public void verifyAllCatogorydrops()
	{
		signup.verfyThirdParty();
	}
	
	@Test(priority=3)
	public void verifyAddTocart()
	{
		boolean addtocart=signup.selectItemfromBestSeller();
		AssertJUnit.assertEquals(addtocart, true);
	}
	
	@Test(priority=4)
	public void verifyItemAddedtoCart()
	{
		boolean itemincart=signup.verifyAddtocart();
		AssertJUnit.assertEquals(itemincart, true);
	}

	@Test(priority=5)
	public void verifyNoOfQuantity() throws InterruptedException
	{
		signup.selectNoOfQuantity();
		
	}
		
	@Test(priority=6)
	public void verifyItemCheckout() throws InterruptedException, AWTException
	{
		signup.verifyCheckout();
		
	}
	
	@Test(priority=7)
	public void verifyShippingAddress() throws InterruptedException, AWTException
	{
		signup.selectAddress();
	}
	
	@Test(priority=8)
	public void verifyShippingOption() throws InterruptedException
	{
		signup.chooseShippingOption();
	}
	
//	@Test(priority=6)
//	public void verifyCardpaymentDetails()
//	{
//		
//		signup.paymentmethod();
//	}
	@Test(priority=9)
	public void verifyOrderPlace() throws InterruptedException, AWTException
	{
		boolean OrderPlaced=signup.placeOrder();
		AssertJUnit.assertEquals(OrderPlaced, true);

	}
	
	@Test(priority=10)
	public void verifyMyOrders() throws InterruptedException
	{
		boolean orderHistory=signup.myOrder();
		AssertJUnit.assertEquals(orderHistory, true);
	}
	
	
	@Test(priority=11)
	public void verifyAccountLoginAndSecurity() throws InterruptedException
	{
		boolean loginsecurity=signup.gotoMyAccountLogin();
		AssertJUnit.assertEquals(loginsecurity, true);
	}
	
	@Test(priority=12)
	public void verifyMyAccountAddresses() throws InterruptedException
	{
		boolean address=signup.gotoMyAddress();
		AssertJUnit.assertEquals(address, true);
	}
	
	@Test(priority=13)
	public void verifyPaymentSourceAvailability() throws InterruptedException
	{
		boolean spdetails=signup.gotoPayment();
		AssertJUnit.assertEquals(spdetails, true);
	}
	
	@Test(priority=14)
	public void verifyAccountSourcePoints() throws InterruptedException
	{
		signup.gotoSourcePoints();
		
	}
	
	@Test(priority=15)
	public void verifyMyWishlist() throws InterruptedException
	{
		boolean wishlist=signup.gotoWishList();
		AssertJUnit.assertEquals(wishlist, true);
	}
	
//	@Test(priority=16)
//	public void verifyUserLogout() throws InterruptedException
//	{
//		boolean logoutmessage=signup.logout();
//		AssertJUnit.assertEquals(logoutmessage, true);
//		
//	}
}
