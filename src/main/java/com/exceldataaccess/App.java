package com.exceldataaccess;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;



/**
 * Hello world!
 *
 */
public class App 
{
	WebDriver driver=null;
	public static int count=0;
	
	public void setup(){
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\ruchira.more\\Downloads\\chromedriver_win32\\chromedriver.exe");
		driver = new ChromeDriver();
	
		//driver=new FirefoxDriver();
		driver.navigate().to("http://60.60.60.251/bugzilla/");
		driver.manage().window().maximize();
		
			
	}
	
	
	@Test(dataProvider="userdata")
  public void Loginscript(String username,String password) throws InterruptedException, IOException, InvalidFormatException{
	  setup();
	  count++;
	  System.out.println("Count out"+count);
	  driver.findElement(By.xpath(".//*[@id='login_link_top']")).click();
	  driver.findElement(By.xpath(".//*[@id='Bugzilla_login_top']")).sendKeys(username);
	  driver.findElement(By.xpath(".//*[@id='Bugzilla_password_top']")).sendKeys(password);
	 driver.findElement(By.xpath(".//*[@id='log_in_top']")).click();
	Thread.sleep(2000);
	 String titleofpage=driver.getTitle();
	 
	 
	 //Assert.assertEquals("Bugzilla Main Page", titleofpage);
	 if(titleofpage.equals("Bugzilla Main Page")){
		 writeXLSXFile("PASS");
		  System.out.println("Count in"+count);
	 }
	 else{
		 writeXLSXFile("Fail");
		  System.out.println("Count in"+count);
	 }
  }
  
  
  @DataProvider(name = "userdata")
	public Object[][] passData() throws IOException {
		ExcelDataConfig config = new ExcelDataConfig(
				"E:\\AngularTutorial\\workspace\\exceldataaccess\\Userlist.xlsx");
		int row = config.getRowCount(0);
		Object[][] data = new Object[row][2];
		for (int i = 0; i < row; i++) {
			data[i][0] = config.getdata(0, i, 0);
			data[i][1] = config.getdata(0, i, 1);
			
		}
		return data;
	}
  
  public static void writeXLSXFile(String status) throws IOException, InvalidFormatException {
		
		String excelFileName = "E:\\AngularTutorial\\workspace\\exceldataaccess\\Userlist.xlsx";//name of excel file
		
		FileInputStream inputStream = new FileInputStream(new File(excelFileName));
      Workbook wb = WorkbookFactory.create(inputStream);

		String sheetName = "Sheet1";//name of sheet

		//Workbook wb = new XSSFWorkbook();
		Sheet sheet = wb.getSheet(sheetName) ;
		
		

		ExcelDataConfig config = new ExcelDataConfig(excelFileName);
		int row1 = config.getRowCount(0);
		
		//iterating r number of rows
		for (int r=0;r <row1 ; r++ )
		{
			Row row = sheet.getRow(r);

				Cell cell = row.createCell(2);
	
	   cell.setCellValue(status);

		}

		FileOutputStream fileOut = new FileOutputStream(excelFileName);

		//write this workbook to an Outputstream.
		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
	}
	

	
	@AfterMethod
	public void close(){
		driver.close();
	}
}



