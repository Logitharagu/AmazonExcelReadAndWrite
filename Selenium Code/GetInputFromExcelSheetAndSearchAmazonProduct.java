package Day19;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.Select;

public class GetInputFromExcelSheetAndSearchAmazonProduct {
	
	public static WebDriver driver;
	public static int sellectBrowser=1;
	public static String fileName="ProductSearch";


	public static void main(String[] args) throws IOException, InterruptedException {
		 browserSellection();
		 browserSetting();
		 getInfoMation();
		 file();
		 exist();
		 
	}
	public static void browserSellection() {
		switch (sellectBrowser) {
		case 1:
			driver=new ChromeDriver();
			System.out.println("User sellect google chrome");
			break;
		case 2:
			driver=new EdgeDriver();
			System.out.println("User sellect microsoft edge");
			break;
		default:
			driver=new ChromeDriver();
			System.out.println("User sellect google chrome");
			break;
		}
	}
	public static void browserSetting() {
		
		driver.manage().window().maximize();
		driver.get("https://www.amazon.in/");
		driver.manage().deleteAllCookies();
		driver.manage().timeouts().pageLoadTimeout(Duration.ofMinutes(20));
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));
	
	}
	public static void getInfoMation() {
	@Nullable
	String title = driver.getTitle();	
	System.out.println("The Title of page: "+title);
	@Nullable
	String currentUrl = driver.getCurrentUrl();	
	System.out.println("page current Url: "+currentUrl);
	String windowHandle = driver.getWindowHandle();
	System.out.println("The Window handle: "+windowHandle);
	}
	public static void file() throws IOException, InterruptedException {
		String Sheet ="Product" ;                  
		String filePath="./Data/"+fileName+".xlsx";
		File Ofile =new File(filePath);
		FileInputStream Oinput=new FileInputStream(Ofile);
		XSSFWorkbook oWorkBook = new XSSFWorkbook(Oinput);
		XSSFSheet sheet = oWorkBook.getSheet(Sheet);
		int lastRowNum = sheet.getLastRowNum();
		for (int iRow = 1; iRow <=lastRowNum; iRow++) {
			XSSFRow row = sheet.getRow(iRow);
			if (row == null) continue;
			XSSFCell productCell = row.getCell(1);
            XSSFCell categoryCell = row.getCell(2);
            if (productCell == null || categoryCell == null) continue;
            String product = productCell.getStringCellValue();
            String category = categoryCell.getStringCellValue();
            System.out.println("the product name: "+product);
            System.out.println("the product name: "+category);
            SearchBox(product, category);
            Thread.sleep(2000);
            PageFoundResult(product, iRow, sheet, oWorkBook, Ofile);
			}
		oWorkBook.close();
		Oinput.close();
		
		}
	 public static void SearchBox(String ProductName,String ProCategories) {
		   WebElement SearchText, Categories,Botton;
		   SearchText=driver.findElement(By.xpath("//input[@id='twotabsearchtextbox']"));
		   SearchText.clear();
		   SearchText.sendKeys(ProductName);
		   
		   Categories=driver.findElement(By.xpath("//select[@id='searchDropdownBox']"));
		   Select Objsel=new Select(Categories);
		   Objsel.selectByVisibleText(ProCategories);
		   
		   Botton=driver.findElement(By.xpath("//input[@id='nav-search-submit-button']"));
		   Botton.click();
	   }
	  public static void PageFoundResult(String result,int rowIndex, XSSFSheet sheet, XSSFWorkbook oWorkBook, File Ofile) throws IOException {
		   WebElement Oresult;
		   Oresult=driver.findElement(By.xpath("(//div[@class='sg-col-inner'][1]//h2)[1]"));
		   //to get the result value as using text
		   String ResultText=Oresult.getText();
		   System.out.println("After Regex: "+ResultText);
		   //next we check the value is greater than zero.you can't directly use because of those are string
		   String Result = ResultText.replaceAll("[^0-9]", "");
		   //still it has string to converted integer
		   int int1 = Integer.parseInt(Result);
		  System.out.println("Extracted Number: " + int1);
		  
		  XSSFRow row = sheet.getRow(rowIndex);
		    if (row == null) row = sheet.createRow(rowIndex);
		    
		    XSSFCell resultCell = row.getCell(3);
		    if (resultCell == null) {
		        resultCell = row.createCell(3);
		    }
		    resultCell.setCellValue(int1);
		    System.out.println("Result written to Excel");
		    
		    FileOutputStream oFileWrite = new FileOutputStream(Ofile);
		    oWorkBook.write(oFileWrite);
		    oFileWrite.flush();
		    oFileWrite.close();
		 
	  }
	  public static void exist() {
		  driver.quit();
	  }
		
		
	}
	


