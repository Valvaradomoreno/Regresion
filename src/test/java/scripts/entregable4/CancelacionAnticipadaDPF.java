package scripts.entregable4;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.MediaEntityBuilder;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.MarkupHelper;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
import com.aventstack.extentreports.reporter.configuration.Theme;
import com.base.web.base.base.ThreadLocalDriver;
import io.cucumber.java.After;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.awt.*;
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
import java.util.concurrent.TimeUnit;


public class CancelacionAnticipadaDPF {

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
	public void CancelacionAnticipadaDPF()throws IOException, InterruptedException, AWTException {

		extent = new ExtentReports();
		spark = new ExtentSparkReporter(System.getProperty("user.dir") + "/test-output/reports4/CancelacionAnticipadaDPF/Report.html");
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
		ArrayList<String> cuenta =readExcelData(2);
		ArrayList<String> razon =readExcelData(3);

		int filas=usuario.size();
  		for(int i=0;i<usuario.size();i++) {
			  try {

  			if(i<(filas)) {

					System.out.println("-----------------------------------");
					System.out.println("Nuevo Test " + i);
					int caso = i+1;
					logger = extent.createTest("Nuevo Test " + caso);

					// ** DESDE AQUI EMPIEZA EL TEST

				driver = new ChromeDriver();
				driver.manage().window().maximize();
				driver.get("https://10.167.21.100:8480/BrowserWebSAD/servlet/BrowserServlet?");

				Thread.sleep(1000);
				driver.findElement(By.id("details-button")).click();
				driver.findElement(By.id("proceed-link")).click();


				WebDriverWait wait = new WebDriverWait(driver, 60);
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
				WebElement iframe1 = driver.findElement(By.xpath("/html/frameset/frame[2]"));
				driver.switchTo().frame(iframe1);

				driver.findElement(By.id("imgError")).click();

				driver.findElement(By.xpath("//img[@alt='Operaciones Minoristas']")).click();

				driver.findElement(By.xpath("//a[contains(text(),'Buscar Depósito ')]")).click();
				driver.switchTo().parentFrame();

				String MainWindow=driver.getWindowHandle();
				Set<String> s1=driver.getWindowHandles();
				Iterator<String> i1=s1.iterator();

				while(i1.hasNext())
				{
					String ChildWindow=i1.next();

					if(!MainWindow.equalsIgnoreCase(ChildWindow))
					{
						driver.switchTo().window(ChildWindow);
					}
				}
				wait.until(ExpectedConditions.elementToBeClickable(By.id("value:1:1:1")));
				driver.findElement(By.id("value:1:1:1")).clear();
				Thread.sleep(200);
				driver.findElement(By.id("value:2:1:1")).clear();
				Thread.sleep(200);
				String attr = driver.findElement(By.xpath("//label[contains(text(),'Número de cuenta')]")).getAttribute("for");
				driver.findElement(By.id(attr)).sendKeys(cuenta.get(i));
				driver.findElement(By.xpath("//a[@alt='Run Selection']")).click();

				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Overview']"))).click();

				String MainWindow2=driver.getWindowHandle();
				Set<String> s2=driver.getWindowHandles();
				Iterator<String> i2=s2.iterator();

				while(i2.hasNext())
				{
					String ChildWindow=i2.next();

					if(!MainWindow2.equalsIgnoreCase(ChildWindow))
					{
						driver.switchTo().window(ChildWindow);
					}
				}

				driver.manage().window().maximize();

				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Cancelacion Anticipada')]"))).click();

				String MainWindow3=driver.getWindowHandle();
				Set<String> s3=driver.getWindowHandles();
				Iterator<String> i3=s3.iterator();

				while(i3.hasNext())
				{
					String ChildWindow=i3.next();

					if(!MainWindow3.equalsIgnoreCase(ChildWindow))
					{
						driver.switchTo().window(ChildWindow);
					}
				}

				wait.until(ExpectedConditions.elementToBeClickable(By.id("radio:tab1:CASH.CANCEL")));
				driver.findElement(By.xpath("/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[2]/td/table/tbody/tr[9]/td[3]/table/tbody/tr/td[2]/input")).click();

				driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();
				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']"))).click();
				Thread.sleep(5000);

				Boolean isPresent = driver.findElements(By.id("errorImg")).size() > 0;
				if (isPresent){
					driver.findElement(By.id("errorImg")).click();
				}else{
					System.out.println("no hay");
				}
				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='messages']/tbody/tr[2]/td[2]/table[2]/tbody/tr/td"))).click();
				String cod1 = driver.findElement(By.xpath("//*[@id='messages']/tbody/tr[2]/td[2]/table[2]/tbody/tr/td")).getText();
				String sSubCadena = cod1.substring(22,39);
				System.out.println(sSubCadena);
				write(i+1, 5, sSubCadena);

				driver.switchTo().window(MainWindow3);
				driver.navigate().refresh();
				Thread.sleep(12000);

				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Run']"))).click();

				String MainWindow4=driver.getWindowHandle();
				Set<String> s4=driver.getWindowHandles();
				Iterator<String> i4=s4.iterator();

				while(i4.hasNext())
				{
					String ChildWindow=i4.next();

					if(!MainWindow4.equalsIgnoreCase(ChildWindow))
					{
						driver.switchTo().window(ChildWindow);
					}
				}

				wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:CLOSURE.REASON"))).click();
				Select selectProducto2 = new Select(driver.findElement(By.id("fieldName:CLOSURE.REASON")));
				selectProducto2.selectByVisibleText(razon.get(i));
				DateFormat dateFormat2 = new SimpleDateFormat("yyyyMMdd");
				String datef = dateFormat2.format(new Date());
				System.out.println(datef);


				driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();
				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']"))).click();
				Thread.sleep(5000);

				String MainWindow8=driver.getWindowHandle();
				Set<String> s8=driver.getWindowHandles();
				Iterator<String> i8=s8.iterator();

				while(i8.hasNext())
				{
					String ChildWindow=i8.next();

					if(!MainWindow8.equalsIgnoreCase(ChildWindow))
					{
						driver.switchTo().window(ChildWindow);
					}
				}
				Thread.sleep(22000);

				driver.findElement(By.xpath("//img[@alt='Details']")).click();
				driver.manage().window().maximize();
				Thread.sleep(2000);

				WebElement iframe0 = driver.findElement(By.xpath("/html/frameset/frameset[1]/frame"));
				driver.switchTo().frame(iframe0);
				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Detalle de liquidacion')]"))).click();
				driver.switchTo().parentFrame();
				Thread.sleep(2000);

				WebElement iframe01 = driver.findElement(By.xpath("/html/frameset/frameset[1]/frameset/frame"));
				driver.switchTo().frame(iframe01);
				driver.findElement(By.xpath("//a[contains(text(),'Ejecutar')]")).click();
				driver.switchTo().parentFrame();

				WebElement iframe02 = driver.findElement(By.xpath("/html/frameset/frameset[2]/frame"));
				driver.switchTo().frame(iframe02);
				Select selectProducto3 = new Select(driver.findElement(By.id("fieldName:CLOSURE.REASON")));
				selectProducto3.selectByVisibleText(razon.get(i));

				driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();
				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']"))).click();
				Thread.sleep(5000);
				if (isPresent){
					driver.findElement(By.id("errorImg")).click();
				}else{
					System.out.println("no hay");
				}
				driver.switchTo().parentFrame();

				String screenshotPath = getScreenShot(driver, "");


				logger.log(Status.PASS, MarkupHelper.createLabel("Final", ExtentColor.GREEN));
				logger.log(Status.PASS,"Final", MediaEntityBuilder.createScreenCaptureFromPath(screenshotPath).build());
				extent.flush();
				write(i+1, 4, "PASSED");

				DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
				String fecha = dateFormat.format(new Date());
				System.out.println(fecha);
				write(i+1, 6, fecha);
					driver.quit();

				}

			}catch (Exception e){

				  String screenshotPath = getScreenShot(driver, "Error");
				  logger.log(Status.FAIL, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Error: "+e, ExtentColor.RED));
				  extent.flush();
				  write(i+1, 4, "FAILED");
				  write(i+1, 5, "");

				  DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
				  String fecha = dateFormat.format(new Date());
				  System.out.println(fecha);
				  write(i+1, 6, fecha);
				  System.out.println("Error: " + e);
				driver.quit();

			}
  		}

}

	public static ArrayList<String> readExcelData(int colNo) throws IOException {

		FileInputStream fis=new FileInputStream(System.getProperty("user.dir") + "/src/Excel/entregable4/CancelacionAnticipadaDPF.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet s=wb.getSheet("CancelacionAnticipadaDPF");
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

	public void write(int i, int celda, String dato) throws IOException {
		String path = System.getProperty("user.dir") + "/src/Excel/entregable4/CancelacionAnticipadaDPF.xlsx";
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

	public static String getScreenShot(WebDriver driver, String screenshotName) throws IOException {
		String dateName = new SimpleDateFormat("yyyyMMddhhmmss").format(new Date());
		TakesScreenshot ts = (TakesScreenshot) driver;
		File source = ts.getScreenshotAs(OutputType.FILE);
		// after execution, you could see a folder "FailedTestsScreenshots" under src folder
		String destination = System.getProperty("user.dir") + "/test-output/reports4/CancelacionAnticipadaDPF/Images/" + screenshotName + dateName + ".png";
		File finalDestination = new File(destination);
		FileUtils.copyFile(source, finalDestination);
		return destination;
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
