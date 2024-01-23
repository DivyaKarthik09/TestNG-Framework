package Excel;

import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import io.github.bonigarcia.wdm.WebDriverManager;

public class ExcelDataProvider 
{
	WebDriver driver;
	@BeforeMethod
	public void setup() 
	{
		System.out.println("Start test");
		WebDriverManager.edgedriver().setup();
		 driver = new EdgeDriver();
		 driver.get("https://www.google.com");
		 driver.manage().window().maximize();
		
	}
	
	@DataProvider(name="excel-data")
	public Object[][] excelDP() {
		Object[][] arrObj = getExcelData("./Excel/TestData.xlsx","DataSet");
		return arrObj;
	}
	public Object[][] getExcelData(String fileName, String sheetName) {
		// TODO Auto-generated method stub
		String[][] data=null;
		try {
			FileInputStream fis = new FileInputStream(fileName);
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sh= wb.getSheet(sheetName);
			XSSFRow row = sh.getRow(0);
			int noOfRows = sh.getPhysicalNumberOfRows();
			int noOfCols = row.getLastCellNum();
			Cell cell;
			data = new String[noOfRows-1][noOfCols];
			
			for(int i=1; i<noOfRows;i++) {
				for(int j=0;j<noOfCols;j++) {
					row =sh.getRow(i);
					cell= row.getCell(j);
					data[i-1][j]=cell.getStringCellValue();
				}
			}
			
		}catch(Exception e)
		{
			System.out.println(e.getMessage());
		}
		return data;
	}@Test(dataProvider="excel-data")
	public void search(String keyWord1, String keyWord2) {
		driver.findElement(By.name("q")).sendKeys(keyWord1);
		driver.findElement(By.name("q")).sendKeys(Keys.ENTER);
	}

}
