package testjava;

import org.testng.annotations.Test;
import org.testng.annotations.Test;
import org.testng.annotations.Test;
import org.testng.annotations.Test;
import org.testng.annotations.Test;

import java.awt.AWTException;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

import org.testng.annotations.Test;

import mainjava.BaseClass;

public class DashboardPageTest extends BaseTest {

	@Test(priority=-3)
	public void verifyDashBoardTest() throws InterruptedException
	{
		dash.thirdPartFigures();
		
	}
	
	@Test(priority=-2)
	public void verifyJapanese() throws InterruptedException
	{
		dash.verifyJapanese();
	
		
	}
	
	@Test(priority=-1)
	public void verifyHasBroProduct() throws InterruptedException
	{
		dash.verifyHasbro();
	
		
	}
	
//	@Test(priority=0)
//	public void verifyVintageProduct() throws InterruptedException
//	{
//		dash.verifyVintage();
//	
//		
//	}
	
	@Test(priority=1)
	public void verifyExclusivesProduct() throws InterruptedException
	{
		dash.verifyExclusives();
	}
	
	@Test(priority=2)
	public void verifySaleProduct() throws InterruptedException
	{
		dash.verifySale();
	}
	
	@Test(priority=3)
	public void verifyCheckOut() throws IOException
	{
		dash.verifyCheckOut();
	}


	@Test(priority=4)
	public void verifychooseShippingOption() throws InterruptedException
	{
		dash.chooseShippingOption();
	}
	
	@Test(priority=5)
	public void verifyPaymentOption() throws InterruptedException, AWTException
	{
		dash.paymentmethod();
	}
	
	@Test(priority=6)
	public void verifyOrder() throws InterruptedException, AWTException
	{
		dash.placeOrder();
	}
	
	
	@Test(priority=7)
	public void verifyingShipngAddress() throws InterruptedException, AWTException, IOException
	{
	dash.shipmyStack();	
	}
	
	@Test(priority=8)
	public void verifyProductSelection() throws InterruptedException, IOException, AWTException
	{
		dash.productSelection();
	}
	
	
	@Test(priority=9)
	public void verifyingPlaceOder() throws InterruptedException
	{
		dash.toPlaceorder();
	}
	
	@Test(priority=10)
	public void verifyOrderDetails() throws InterruptedException, IOException
	{
		dash.orderDetails();
	}
	
	
	@Test(priority=11)
	public void verifyproductSelection() throws InterruptedException, AWTException, IOException
	{
		dash.repetation();
	}
	
	@Test(priority=13)
	public void logout()
	{
		dash.logout();
	}
	
}
