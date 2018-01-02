package Demo1.Excelforjenkins;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

public class Gmail {
	
	XSSFWorkbook book;
	XSSFSheet sheet;
	XSSFRow column;
	int totalrows,totalcolumns;
	public String[][] arr=new String[totalrows][totalcolumns];
	int row,col;
	WebDriver driver;	
 
	public String URL="https://accounts.google.com/signin/v2/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F%3Ftab%3Dwm&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=ServiceLogin";
	
	@Test
	public void loadingexcelsheet() throws Exception
	{
		String Response="Fail";
		
			try {
			File fs=new File("C:\\Users\\vikram\\Desktop\\gmail.xlsx");
			FileInputStream file=new FileInputStream(fs);
			book=new XSSFWorkbook(file);
			Response="Pass";
			} 	catch (Exception e){
				Response=Response+":"+e.getMessage();
				System.out.println("upload status==" +Response);
				e.printStackTrace();
			} //end catch statement
		
	
		sheet=book.getSheetAt(0);
		column=sheet.getRow(0);
		
		totalrows=sheet.getLastRowNum()+1;
		totalcolumns=column.getLastCellNum();
		
		arr=new String[totalrows][totalcolumns];
		
		for(row=0;row<totalrows;row++){
			for(col=0;col<totalcolumns;col++)
			{
				arr[row][col]=sheet.getRow(row).getCell(col).getStringCellValue();
				System.out.println("From excel== " + arr[row][col]);
				
			}
		}

		
		for(row=1;row<totalrows;row++){
			
			int colapplication=0;
			int colusername=1;
			int colpassword=2;
			
			
			
			String values=arr[row][colapplication];
			System.out.println("data from excel == " + arr[row][colusername] +"== "  + arr[row][colpassword]);
			
			
					driver=new FirefoxDriver();
					driver.get(URL);
					String response="Failed; It seems user ID or password failed"+arr[row][colusername] +"== "  + arr[row][colpassword];
					
					try
					{
					
					/*String currentURl=drivergetCurrentUrl();
					Assert.assertTrue(currentURl.contains("google.com"));
					*/
					WebElement Username1=driver.findElement(By.xpath("//input[@id='identifierId' and @name='identifier']"));
					Username1.sendKeys(arr[row][colusername]);
					
					WebElement Nexttab=driver.findElement(By.xpath("//span[@class='RveJvd snByac']"));
					Nexttab.click();
					
					Thread.sleep(10000);
					
					driver.findElement(By.xpath("//input[@type='password' and @name='password']")).sendKeys(arr[row][colpassword]);
					
					driver.findElement(By.xpath("//span[@class='RveJvd snByac']")).click();
					
					WebDriverWait wait=new WebDriverWait(driver, 50);
					wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[@class='gb_ab gbii']")));
					
					
					driver.findElement(By.xpath("//span[@class='gb_ab gbii']")).click();
					
					driver.findElement(By.xpath("//a[@id='gb_71']")).click();
					
					response= arr[row][colusername] +"== "  + arr[row][colpassword];
					System.out.println("Used ID ==" +response);
					driver.close();
					}catch(Exception e)
					{
						response=response+ " : " + e.getMessage();
						System.out.println("Error ==" +response);
						e.printStackTrace();
						
					}
					
			
			}
			
				
	 } 
	}
