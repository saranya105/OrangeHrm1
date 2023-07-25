package com.orange;

import java.io.FileInputStream;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class MyInfo {

	public static void main(String[] args) {
try {
	    	ChromeDriver driver =new ChromeDriver();
		    driver.manage().window().maximize();
		    driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		    driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login");
		    FileInputStream file = new FileInputStream("C:\\selenium\\orange.xlsx");
		    XSSFWorkbook workbook = new XSSFWorkbook(file);
		    XSSFSheet sheet = workbook.getSheet("Sheet1");
		    int rowcount = sheet.getLastRowNum();
		    int colcount = sheet.getRow(1).getLastCellNum();
		    ArrayList allValues=new ArrayList();
		    for(Row row : sheet) {
		    	Cell cell=row.getCell(1);
		    	allValues.add(cell);
		    }
		    for(int i=1; i<=rowcount; i++)
		    {
		    	String MyInfo= allValues.get(0).toString();
		    	String OldId= allValues.get(1).toString();
		    	String Input= allValues.get(2).toString();
		    	String DriverLicenseNumber= allValues.get(3).toString();
		    	String LicenseNumber= allValues.get(4).toString();
		    	String LicenseExpiryDate= allValues.get(5).toString();
		    	String ExpiryDate= allValues.get(6).toString();
		    	String SSNNumber= allValues.get(7).toString();
		    	String Number= allValues.get(8).toString();
		    	String SINNumber= allValues.get(9).toString();
		    	String SINumber= allValues.get(10).toString();
		    	String BloodType= allValues.get(11).toString();
		    	String BloodTypes= allValues.get(12).toString();
		    	String Type= allValues.get(13).toString();
		    	String MilitaryService= allValues.get(14).toString();
		    	String Service= allValues.get(15).toString();
		    	String Smoker= allValues.get(16).toString();
		    	String save= allValues.get(17).toString();
		    	String save1= allValues.get(18).toString();
		    	String ContactDetails= allValues.get(19).toString();
		    	String Street1= allValues.get(20).toString();
		    	String ad1= allValues.get(21).toString();
		    	String City= allValues.get(22).toString();
		    	String City1= allValues.get(23).toString();
		    	String State= allValues.get(24).toString();
		    	String Statename= allValues.get(25).toString();
		    	String PostalCode= allValues.get(27).toString();
		    	String Code= allValues.get(28).toString();
		    	String Home= allValues.get(29).toString();
		    	String Hno1= allValues.get(30).toString();
		    	String Mobile= allValues.get(31).toString();
		    	String Mno1= allValues.get(32).toString();
		    	String OtherEmail= allValues.get(33).toString();
		    	String Email= allValues.get(34).toString();
		    	String save2= allValues.get(35).toString();
		    	
		        driver.findElement(By.name("username")).sendKeys("Admin");      
		        Thread.sleep(500);
		        driver.findElement(By.name("password")).sendKeys("admin123");
		        driver.findElement(By.xpath("//button[@type='submit']")).click();
		        Thread.sleep(500);
		        
		        
		  //Personal Details
		        driver.findElement(By.xpath(MyInfo)).click();
		        Thread.sleep(1000);
		        driver.findElement(By.xpath(OldId)).sendKeys(Input);
		        Thread.sleep(1000);
		        driver.findElement(By.xpath(DriverLicenseNumber)).sendKeys(LicenseNumber);
		        Thread.sleep(1000);
		        driver.findElement(By.xpath(LicenseExpiryDate)).sendKeys(ExpiryDate);
		        Thread.sleep(1000);
		        driver.findElement(By.xpath(SSNNumber)).sendKeys(Number);
		        Thread.sleep(1000);
		        driver.findElement(By.xpath(SINNumber)).sendKeys(SINumber);
		        Thread.sleep(1000);
		        driver.findElement(By.xpath(MilitaryService)).sendKeys(Service);
		        Thread.sleep(1000);
		        driver.findElement(By.xpath(Smoker)).click();
		        Thread.sleep(1000);
		        driver.findElement(By.xpath(save)).click();
		        Thread.sleep(1000);
		        driver.findElement(By.xpath(BloodType)).click();
		        Thread.sleep(1000);
		       List<WebElement> list = driver.findElements(By.xpath(BloodTypes));
				 Iterator<WebElement> iterator = list.iterator();
				 while (iterator.hasNext()) {
					WebElement webElement = (WebElement) iterator.next();
					if(webElement.getText().equals(Type)) {
						webElement.click();
					}
				 }
				 Thread.sleep(1000);
				  driver.findElement(By.xpath(save1)).click();
				 
		//Contact Details
	 
				 driver.findElement(By.xpath(ContactDetails)).click();
				 Thread.sleep(1000);
				 driver.findElement(By.xpath(Street1)).sendKeys(ad1);
				 driver.findElement(By.xpath(City)).sendKeys(City1);
				 Thread.sleep(1000);
				 driver.findElement(By.xpath(State)).sendKeys(Statename);
				 driver.findElement(By.xpath(PostalCode)).sendKeys(Code);
			     driver.findElement(By.xpath(Home)).sendKeys(Hno1);		
			     Thread.sleep(1000);
				 driver.findElement(By.xpath(Mobile)).sendKeys(Mno1);
				 Thread.sleep(1000);
				 driver.findElement(By.xpath(OtherEmail)).sendKeys(Email);
				 driver.findElement(By.xpath(save2)).click();
				 
				 Thread.sleep(1000);
				 driver.navigate().back();
		    }
			   driver.quit();
			    workbook.close();
	     	} 
		    catch (Exception e) {
			   System.out.println(e.getMessage());
	     	}
	  }
	}
