package votw;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Test;

public class votwLoginPage {
	
	public static WebDriver driver;
	public static String BaseUrl="https://www.facebook.com/";
	@BeforeSuite
	public void getBrowser(){
		System.setProperty("webdriver.chrome.driver", "E:\\Imp java Stuff\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.get(BaseUrl);
		driver.manage().window().maximize();
	}
	
	
	@Test
	public void testMethod(){
		System.out.println("Into the method");
	}
	
	@Test
	public void testMethod1(){
		System.out.println("Into the method1");
	}
	
	@Test
	public void testMethod2(){
		System.out.println("Into the method1");
	}
	
	@Test
	public void testMethod3(){
		System.out.println("Into the method1123");
	}
	@AfterSuite
	public void closeBrowser(){
		driver.quit();
	}

}
