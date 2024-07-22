package hardcore;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Properties;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Script1 {

	public static void main(String[] args) throws IOException, InterruptedException {
		FileInputStream fis=new FileInputStream("./src/test/resources/commondata.properties");
		Properties p=new Properties();
		p.load(fis);
		String url=p.getProperty("url");
		String Username=p.getProperty("username");
		String Password=p.getProperty("password");
		WebDriver driver= new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
		driver.get(url);
		driver.findElement(By.name("user_name")).sendKeys(Username);
		driver.findElement(By.name("user_password")).sendKeys(Password);
		driver.findElement(By.id("submitButton")).click();
		Thread.sleep(3000);
		driver.findElement(By.xpath("//img[@src='themes/softed/images/user.PNG']")).click();
		driver.findElement(By.linkText("Sign Out")).click();
		Thread.sleep(3000);
		driver.quit();
		
		
		
		
		
		
		
		
	}

}
