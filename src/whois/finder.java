package whois;

import java.io.File;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;


import java.io.IOException;
import java.util.Locale;

import java.io.File;
import java.io.IOException;
import jxl.*;
import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.read.biff.BiffException;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

//import com.gargoylesoftware.htmlunit.javascript.host.Console;

public class finder {
	public static String domainName;
	
	public static String site = "http://www.tradeindia.com/";
	private static WritableCellFormat timesBoldUnderline;
	private static WritableCellFormat times;
	
	private static void createLabel(WritableSheet sheet)
		      throws WriteException {
		    // Lets create a times font
		    WritableFont times10pt = new WritableFont(WritableFont.TIMES, 10);
		    // Define the cell format
		    times = new WritableCellFormat(times10pt);
		    // Lets automatically wrap the cells
		    times.setWrap(true);

		    // create create a bold font with underlines
		    WritableFont times10ptBoldUnderline = new WritableFont(WritableFont.TIMES, 10, WritableFont.BOLD, false,
		        UnderlineStyle.SINGLE);
		    timesBoldUnderline = new WritableCellFormat(times10ptBoldUnderline);
		    // Lets automatically wrap the cells
		    timesBoldUnderline.setWrap(true);

		    CellView cv = new CellView();
		    cv.setFormat(times);
		    cv.setFormat(timesBoldUnderline);
		    cv.setAutosize(true);

		    // Write a few headers
		    addCaption(sheet, 0, 0, "Company Name");
		    addCaption(sheet, 1, 0, "Key Personnel");
		    addCaption(sheet, 2, 0, "Mobile");
		    addCaption(sheet, 3, 0, "Phone");
		  }
	private static void addCaption(WritableSheet sheet, int column, int row, String s)
		      throws RowsExceededException, WriteException {
		    Label label;
		    label = new Label(column, row, s, timesBoldUnderline);
		    sheet.addCell(label);
		  }
	
	
	public static void main(String[] args){
		
		File outFile = new File("./output.xls");
	    
	    WorkbookSettings wbSettings = new WorkbookSettings();
	    
	    wbSettings.setLocale(new Locale("en", "EN"));
	    
	    WritableWorkbook workbook;
		
	    try {
			workbook = Workbook.createWorkbook(outFile, wbSettings);
		
	    
		    workbook.createSheet("Report", 0);
		    
			WritableSheet excelSheet = workbook.getSheet(0);
			createLabel(excelSheet);
		
		
		
		
			try{
				
				System.out.println("finder tool start");
				
				//webdriver
				File file = new File("./chromedriver.exe");
				
			    System.setProperty("webdriver.chrome.driver", file.getAbsolutePath());
			    
			    WebDriver driver = new ChromeDriver();
			    
			    driver.get(site);
			    
				//-webdriver
			    
			    WebElement element = driver.findElement(By.xpath("//*[@id='advanced-search']/div[1]"));
			    element.click();
			    WebElement products = driver.findElement(By.xpath("//*[@id='advanced-search']/div[2]/div[2]/form/div[1]/div[1]/select/option[3]"));
			    products.click();
			    WebElement keywords = driver.findElement(By.id("Products"));
			    keywords.sendKeys("automobile");
			    WebElement city = driver.findElement(By.name("cities"));
			    city.sendKeys("Gurgaon");
			    
			    List<WebElement> exporter = driver.findElements(By.id("Exporter"));
			    for(int i=0; i<exporter.size(); i++){
			    	exporter.get(i).click();
		        }
			    WebElement searchbtn = driver.findElement(By.name("submit"));
			    searchbtn.click();
			    Thread.sleep(1000);
			    //open output excel
			    
			    // workbook create and open moved from here
			    
				int flag=1;
				int j=0;
				int count=0;
				
				String nextPageBase = "/html/body/div[5]/div/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/table/tbody/tr/td/form/table[4]/tbody/tr[1]/td/span[2]/";
				
				while(flag==1)
				{ //multiple pages
					
					String resultString;
					String viewDetails;
					
					for(int it=1; it<30; it++){ //multiple results on same page
						System.out.println("for loop it= "+it);
						resultString = "/html/body/div[5]/div/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/table/tbody/tr/td/form/table[2]/tbody/tr["+it+"]/td[1]/span[1]/a/font";
						viewDetails = "/html/body/div[5]/div/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[5]/td/table/tbody/tr/td/form/table[2]/tbody/tr["+it+"]/td[1]/span[2]/table[2]/tbody/tr[2]/td[5]/a";
						
						if(driver.findElements(By.xpath(resultString)).size() == 0) //xpath not found in combination
							continue;
						else
						{	//xpath found in combination
							WebElement res = driver.findElement(By.xpath(resultString));
							String COMPANY = res.getText();
							
							count++;
							System.out.println(count+":  "+COMPANY);
							
							WebElement details = driver.findElement(By.xpath(viewDetails));
							Thread.sleep(1000);
						    
							//res.click();
							try{
								details.click();								
							}
							catch(Exception e){
								continue;
							}
							
						    //System.out.println("click pt");
						    Thread.sleep(1000);
						    //Store the current window handle
						    
						    //Perform the click operation that opens new window
						    //WebElement contactdetails = driver.findElement(By.xpath("/html/body/div[4]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[5]/td/a"));
							
							if(driver.findElements(By.linkText("View Contact Details")).size()!=0)
							{
								
								WebElement contactdetails = driver.findElement(By.linkText("View Contact Details"));
								String winHandleBefore = driver.getWindowHandle();
								contactdetails.click();
							    Thread.sleep(1000);
							    //Switch to new window opened
							    for(String winHandle : driver.getWindowHandles()){
							        driver.switchTo().window(winHandle);
							    }
							    
							    //perform action on new window
							    WebElement contactName = driver.findElement(By.xpath("/html/body/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td[3]"));
							    String CONTACT = contactName.getText();
							    //System.out.println(CONTACT);
							    
							    String MOB;
							    if(driver.findElements(By.xpath("/html/body/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]")).size()!=0)
							    {
							    	WebElement mobile = driver.findElement(By.xpath("/html/body/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[3]"));
								    MOB = mobile.getText();							    	
							    }
							    else
							    {
							    	MOB = " ";
							    }
							    
							    //System.out.println(MOB);
							    
							    
							    if(driver.findElements(By.xpath("/html/body/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td[3]")).size()!=0)
							    {
							    	WebElement phone = driver.findElement(By.xpath("/html/body/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td[3]"));
								    String PH = phone.getText();
								    
								  //WRITE TO EXCEL START
								    j++;
								    
								    for(int i = 0; i < 4; i++) {
					
							      		Label label;
							      		if(i==0)
							      		{
							      			label = new Label(i,j,COMPANY,times);
							      			excelSheet.addCell(label);
							      		}
							      		
							      		if(i==1)
							      		{
							      			label = new Label(i,j,CONTACT,times);
							      			excelSheet.addCell(label);
							      		}
							      		if(i==2)
							      		{
							      			label = new Label(i,j,MOB,times);
							      			excelSheet.addCell(label);
							      		}
							      		if(i==3)
							      		{
							      			label = new Label(i,j,PH,times);
							      			excelSheet.addCell(label);
							      		}
							      	} //WRITE TO EXCEL END
							    }
							    else
							    {
							    	
							    	//WRITE TO EXCEL START
								    j++;
								    
								    for(int i = 0; i < 4; i++) {
					
							      		Label label;
							      		if(i==0)
							      		{
							      			label = new Label(i,j,COMPANY,times);
							      			excelSheet.addCell(label);
							      		}
							      		
							      		if(i==1)
							      		{
							      			label = new Label(i,j,CONTACT,times);
							      			excelSheet.addCell(label);
							      		}
							      		if(i==2)
							      		{
							      			label = new Label(i,j,MOB,times);
							      			excelSheet.addCell(label);
							      		}
							      	} //WRITE TO EXCEL END
							    }
							    
							    
							    //Close the new window, if that window no more required
							    driver.close();
							    
							    driver.switchTo().window(winHandleBefore);
							    
							  
							    
							} // contact details if end
							
							else
							{	
								
								if(driver.findElements(By.xpath("/html/body/div[4]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[5]/td/div")).size()!=0)
						    	{
									WebElement mobile = driver.findElement(By.xpath("/html/body/div[4]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[5]/td/div"));
									String MOB = mobile.getText();
									WebElement contact1 = driver.findElement(By.xpath("/html/body/div[4]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]"));
									WebElement contact2 = driver.findElement(By.xpath("/html/body/div[4]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td[2]"));
									String CONTACT =  contact1.getText() + contact2.getText();
									
									// WRITE TO EXCEL START
									j++;
								    
								    for(int i = 0; i < 4; i++) {
					
							      		Label label;
							      		if(i==0)
							      		{
							      			label = new Label(i,j,COMPANY,times);
							      			excelSheet.addCell(label);
							      		}
							      		
							      		if(i==1)
							      		{
							      			label = new Label(i,j,CONTACT,times);
							      			excelSheet.addCell(label);
							      		}
							      		if(i==2)
							      		{
							      			label = new Label(i,j,MOB,times);
							      			excelSheet.addCell(label);
							      		}
							      	} //WRITE TO EXCEL END
									
						    	}
						    	else
						    	{
						    		if(driver.findElements(By.xpath("/html/body/div[4]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td/div")).size()!=0)
						    		{
						    			WebElement mobile = driver.findElement(By.xpath("/html/body/div[4]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td/div"));
										String MOB = mobile.getText();
										WebElement contact1 = driver.findElement(By.xpath("/html/body/div[4]/div/table/tbody/tr[3]/td/table/tbody/tr/td[2]/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]"));
										String CONTACT =  contact1.getText();
										
										// WRITE TO EXCEL START
										
										j++;
									    
									    for(int i = 0; i < 4; i++) {
						
								      		Label label;
								      		if(i==0)
								      		{
								      			label = new Label(i,j,COMPANY,times);
								      			excelSheet.addCell(label);
								      		}
								      		
								      		if(i==1)
								      		{
								      			label = new Label(i,j,CONTACT,times);
								      			excelSheet.addCell(label);
								      		}
								      		if(i==2)
								      		{
								      			label = new Label(i,j,MOB,times);
								      			excelSheet.addCell(label);
								      		}
								      	} //WRITE TO EXCEL END
						    		}
						    		else
						    		{
						    			// WRITE TO EXCEL START
										
										j++;
									    
									    for(int i = 0; i < 4; i++) {
						
								      		Label label;
								      		if(i==0)
								      		{
								      			label = new Label(i,j,COMPANY,times);
								      			excelSheet.addCell(label);
								      		}
								      		
								      	} //WRITE TO EXCEL END
						    		}
						    		
									
						    	}
								
								
							    
							}    //contact details else end
								
						    
							    
						    driver.navigate().back();
						    Thread.sleep(1000);
						} //if-else end
	
					} //for end
					 
					if(driver.findElements(By.xpath(nextPageBase+"span/a")).size()==0)
					{
						System.out.println("cant find next page, exiting");
						flag=0;
					}
					else
					{
						nextPageBase = nextPageBase+"span/";
						System.out.println(nextPageBase);
						WebElement nextPage = driver.findElement(By.xpath(nextPageBase+"a"));
						System.out.println("next page else before click");
						
						//driver.navigate().refresh();
						JavascriptExecutor jse = (JavascriptExecutor)driver;
						jse.executeScript("window.scrollBy(0,30000)", "");
						Thread.sleep(1000);
						
						nextPage.click();
						Thread.sleep(1000);
						System.out.println("clicking next page");
					}
				} //while end
				
				workbook.write();
				
				workbook.close();
			
				driver.quit();
			
			} //end of try
			catch(Exception e)
			{
				e.printStackTrace();
				System.out.println("Exception");
				
				workbook.write();
				
				workbook.close();
			}
		
	    } catch (IOException | WriteException e1) {
			
			e1.printStackTrace();
			System.out.println("Workbook opening error");
		}
	} //main end
} //finder class end
