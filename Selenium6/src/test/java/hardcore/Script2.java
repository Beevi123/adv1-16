package hardcore;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;
import java.util.Properties;
import java.util.Random;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class Script2 {

	public static void main(String[] args) throws InterruptedException, IOException {
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
		driver.findElement(By.linkText("Organizations")).click();
		driver.findElement(By.xpath("//img[@src='themes/softed/images/btnL3Add.gif']")).click();
		Random random=new Random();
		int ranNum=random.nextInt(100);
		FileInputStream fis1=new FileInputStream("C:\\Users\\faizh\\OneDrive\\InsertData.xlsx");
		Workbook wb=WorkbookFactory.create(fis1);
	String data=wb.getSheet("Sheet3").getRow(0).getCell(0).toString()+ranNum;
	System.out.println(data);
	driver.findElement(By.name("accountname")).sendKeys(data);
	Thread.sleep(3000);
	driver.findElement(By.id("phone")).sendKeys("123456");
	Thread.sleep(3000);

	Select sd=new Select(driver.findElement(By.name("industry")));
List<WebElement> alloptions=	sd.getOptions();

System.out.println(alloptions.size());
for(int i=0;i<alloptions.size();i++) {
	System.out.println(alloptions.get(i).getText());
}
	
	
		
		

	}

}
