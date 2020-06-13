package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Excel {

	public static void main(String[] args) throws Throwable {
		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\prabhu\\eclipse-workspace\\ExcelDatapass\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();

		driver.get("http://newtours.demoaut.com/");

		File f = new File("C:\\Users\\prabhu\\eclipse-workspace\\ExcelDatapass\\ExcelDatapass.xlsx");
		FileInputStream stream=new FileInputStream(f);
		Workbook w=new XSSFWorkbook(stream);
		Sheet s = w.getSheet("Sheet1");
		for (int i = 1; i <s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(i);
			for (int j = 0; j <r.getPhysicalNumberOfCells(); j++) {
				Cell c = r.getCell(j);
			DataFormatter format=new DataFormatter();
			String celltype = format.formatCellValue(c);
			if(j==0) {
			driver.findElement(By.name("userName")).sendKeys(celltype);
			}
			else  {
			driver.findElement(By.name("password")).sendKeys(celltype);
			
			
			}
			
			
		}
			driver.findElement(By.name("login")).click();
			Thread.sleep(2000);
			driver.findElement(By.linkText("Home")).click();
			String title = driver.getTitle();
			Cell createcell = r.createCell(2);
			createcell.setCellValue(title);
			
			
		FileOutputStream o=new FileOutputStream(f);
		w.write(o);
		
	}

}
}