package scripts;

import java.awt.AWTException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.Set;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.MarkupHelper;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;
import com.base.web.base.base.Excel;
import com.base.web.base.base.ThreadLocalDriver;
import com.base.web.base.base.writeExcel;
import io.cucumber.java.After;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;


public class AltaCuentaAhorro {

    WebDriver driver;
	public ExtentSparkReporter spark;
	public ExtentReports extent;
	public ExtentTest logger;
	@BeforeTest
    public void startTest() {

    }
        
	@BeforeMethod
    public void openApplication() {
    	System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir") + "/src/main/resources/drivers/chromedriver");

	}


	@Test
	public void altaCuentaAhorro()throws IOException, InterruptedException, AWTException {

		extent = new ExtentReports();
		spark = new ExtentSparkReporter(System.getProperty("user.dir") + "/test-output/reports/AltaCuentaAhorro/Report.html");
		extent.attachReporter(spark);
		extent.setSystemInfo("Host Name", "SoftwareTestingMaterial");
		extent.setSystemInfo("Environment", "Production");
		extent.setSystemInfo("User Name", "Rajkumar SM");
		spark.config().setDocumentTitle("Title of the Report Comes here ");
		spark.config().setReportName("Name of the Report Comes here ");
		spark.config().setTheme(Theme.STANDARD);


    	Thread.sleep(1500);

		ArrayList<String> usuario=readExcelData(0);
		ArrayList<String> contraseña =readExcelData(1);
		ArrayList<String> documento =readExcelData(2);
		ArrayList<String> ejecutivo =readExcelData(3);

		int filas=usuario.size();
  		for(int i=0;i<usuario.size();i++) {
			try {
  			if(i<(filas)) {


					System.out.println("-----------------------------------");
					System.out.println("Nuevo Test " + i);
					int caso = i+1;
					logger = extent.createTest("Nuevo Test " + caso);

					driver = new ChromeDriver();
					driver.manage().window().maximize();
					driver.get("https://10.167.21.100:8480/BrowserWebSAD/servlet/BrowserServlet?");

					Thread.sleep(1000);
					driver.findElement(By.id("details-button")).click();
					driver.findElement(By.id("proceed-link")).click();

					WebDriverWait wait = new WebDriverWait(driver, 40);
					wait.until(ExpectedConditions.elementToBeClickable(By.id("signOnName")));
					driver.findElement(By.id("signOnName")).sendKeys(usuario.get(i));
					driver.findElement(By.id("password")).sendKeys(contraseña.get(i));
					driver.findElement(By.id("sign-in")).click();

					WebElement iframe = driver.findElement(By.xpath("/html/frameset/frame[1]"));
					driver.switchTo().frame(iframe);

					Thread.sleep(2000);
					String exp_message = "Sign Off";
					String actual = driver.findElement(By.xpath("//a[contains(text(),'Sign Off')]")).getText();
					Assert.assertEquals(exp_message, actual);
					System.out.println("assert complete");
					driver.switchTo().parentFrame();

					Thread.sleep(1000);
					WebElement iframe2 = driver.findElement(By.xpath("/html/frameset/frame[2]"));
					driver.switchTo().frame(iframe2);
					driver.findElement(By.id("imgError")).click();

					driver.findElement(By.xpath("//a[contains(text(),'Catálogo de Productos ')]")).click();
					driver.switchTo().parentFrame();

					String MainWindow = driver.getWindowHandle();
					Set<String> s1 = driver.getWindowHandles();
					Iterator<String> i1 = s1.iterator();

					while (i1.hasNext()) {
						String ChildWindow = i1.next();

						if (!MainWindow.equalsIgnoreCase(ChildWindow)) {
							driver.switchTo().window(ChildWindow);
						}
					}
					driver.manage().window().maximize();
					WebElement iframe3 = driver.findElement(By.xpath("/html/frameset/frameset[2]/frameset[1]/frame[2]"));
					driver.switchTo().frame(iframe3);
					driver.findElement(By.id("treestop1")).click();
					driver.findElement(By.xpath("//*[@id='r1']/td[4]/a/img")).click();
					driver.switchTo().parentFrame();

					WebElement iframe4 = driver.findElement(By.xpath("/html/frameset/frameset[2]/frameset[2]/frame[2]"));
					driver.switchTo().frame(iframe4);
					driver.findElement(By.xpath("//*[@id='r1']/td[3]/a/img")).click();
					driver.switchTo().parentFrame();

					String MainWindow2 = driver.getWindowHandle();
					Set<String> s2 = driver.getWindowHandles();
					Iterator<String> i2 = s2.iterator();

					while (i2.hasNext()) {
						String ChildWindow = i2.next();

						if (!MainWindow2.equalsIgnoreCase(ChildWindow)) {
							driver.switchTo().window(ChildWindow);
						}
					}

					driver.manage().window().maximize();
					driver.findElement(By.id("fieldName:CUSTOMER:1")).sendKeys(documento.get(i));
					driver.findElement(By.id("fieldName:CURRENCY")).sendKeys("PEN");

					driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();

					wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:PRIMARY.OFFICER")));
					driver.findElement(By.id("fieldName:PRIMARY.OFFICER")).sendKeys(ejecutivo.get(i));
					driver.findElement(By.xpath("/html/body/div[5]/fieldset[4]/div/div/form[1]/div[3]/table/tbody/tr[3]/td/table[1]/tbody/tr[11]/td[3]/table/tbody/tr/td[2]/input")).click();
					driver.findElement(By.xpath("/html/body/div[5]/fieldset[4]/div/div/form[1]/div[3]/table/tbody/tr[3]/td/table[1]/tbody/tr[12]/td[3]/table/tbody/tr/td[2]/input")).click();

					String cod = driver.findElement(By.id("disabled_ACCOUNT.REFERENCE")).getText();
					System.out.println("CUENTA : " + cod);
					String arreglo = driver.findElement(By.id("disabled_ARRANGEMENT")).getText();
					System.out.println("ARREGLO : " + arreglo);

					String screenshotPath = getScreenShot(driver, "Datos agregados");


					driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();

					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']")));
					driver.findElement(By.xpath("//img[@alt='Commit the deal']")).click();

					String screenshotPath1 = getScreenShot(driver, "Commit");


					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div[3]/div/div[2]/form[1]/div[3]/table/tbody/tr[2]/td/table/tbody/tr/td[3]/select")));

					Select selectProducto = new Select(driver.findElement(By.xpath("/html/body/div[3]/div/div[2]/form[1]/div[3]/table/tbody/tr[2]/td/table/tbody/tr/td[3]/select")));
					selectProducto.selectByVisibleText("RECEIVED");
					driver.findElement(By.id("errorImg")).click();


					String cod2 = driver.findElement(By.xpath("//*[@id='messages']/tbody/tr[2]/td[2]/table[2]/tbody/tr/td")).getText();
					String sSubCadena = cod2.substring(22,39);
					System.out.println(sSubCadena);

					write(i+1, 5, sSubCadena);

					DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
					String date = dateFormat.format(new Date());
					System.out.println(date);
					write(i+1, 6, date);

					String screenshotPath2 = getScreenShot(driver, "ALTA CREADA");

					logger.log(Status.PASS, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Datos Agregados", ExtentColor.GREEN));
					logger.log(Status.PASS, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath1) + " Commit", ExtentColor.GREEN));
					logger.log(Status.PASS, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath2) + " Alta creada", ExtentColor.GREEN));

					extent.flush();
					write(i+1, 4, "PASSED");
					driver.quit();
				}

			}catch (Exception e){
				String screenshotPath = getScreenShot(driver, "ERROR");
				logger.log(Status.FAIL, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Test Case FAIL: " +e, ExtentColor.RED));
				extent.flush();
				write(i+1, 5, "");
				write(i+1, 4, "FAILED");

				DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
				String date = dateFormat.format(new Date());
				write(i+1, 6, date);
				driver.quit();
			}
  		}

}

	public void write(int i, int celda, String dato) throws IOException {
		String path = System.getProperty("user.dir") + "/src/Excel/AltaCuentaAhorro.xlsx";
		FileInputStream fs = new FileInputStream(path);
		Workbook wb = new XSSFWorkbook(fs);
		Sheet sheet1 = wb.getSheetAt(0);
		int lastRow = sheet1.getLastRowNum();
		//for(int i=0; i<=lastRow; i++){
		Row row = sheet1.getRow(i);
		Cell cell = row.createCell(celda);

		cell.setCellValue(dato);

		//}

		FileOutputStream fos = new FileOutputStream(path);
		wb.write(fos);
		fos.close();
	}
	//This method is to capture the screenshot and return the path of the screenshot.
	public static String getScreenShot(WebDriver driver, String screenshotName) throws IOException {
		String dateName = new SimpleDateFormat("yyyyMMddhhmmss").format(new Date());
		TakesScreenshot ts = (TakesScreenshot) driver;
		File source = ts.getScreenshotAs(OutputType.FILE);
// after execution, you could see a folder "FailedTestsScreenshots" under src folder
		String destination = System.getProperty("user.dir") + "/reports/AltaCuentaAhorro/Images/" + screenshotName + dateName + ".png";
		File finalDestination = new File(destination);
		FileUtils.copyFile(source, finalDestination);
		return destination;
	}

	public static ArrayList<String> readExcelData(int colNo) throws IOException {
		
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir") + "/src/Excel/AltaCuentaAhorro.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet s=wb.getSheet("AltaCuentaAhorro");
		Iterator<Row> rowIterator=s.iterator();
		rowIterator.next();
		//rowIterator.next();
		ArrayList<String> list=new ArrayList<String>();
		while(rowIterator.hasNext()) {
			list.add(rowIterator.next().getCell(colNo).getStringCellValue());
		}
		System.out.println("List: "+list);
		return list;
	}




	@After
	public void tearDown() throws Exception
	{
		driver = ThreadLocalDriver.getTLWebDriver();
		if(driver != null){
			driver.quit();
		}
	}
	
	
}
