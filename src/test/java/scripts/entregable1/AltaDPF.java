package scripts.entregable1;

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
import java.time.Duration;


public class AltaDPF {

    WebDriver driver = ThreadLocalDriver.getTLWebDriver();
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
	public void AltaDPF()throws IOException, InterruptedException, AWTException {

		extent = new ExtentReports();
		spark = new ExtentSparkReporter(System.getProperty("user.dir") + "/test-output/reports/AltaDPF/Report.html");
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
		ArrayList<String> monto =readExcelData(4);
		ArrayList<String> plazo =readExcelData(5);
		ArrayList<String> cuenta =readExcelData(6);

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

				//WebDriverWait wait = new WebDriverWait(driver, 30);
				//wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//a[contains(text(),'Sign Off')]")));

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

				driver.findElement(By.xpath("//a[contains(text(),'Catálogo de Productos ')]")).click();
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
				driver.manage().window().maximize();
				WebElement iframe2 = driver.findElement(By.xpath("/html/frameset/frameset[2]/frameset[1]/frame[2]"));
				driver.switchTo().frame(iframe2);
				driver.findElement(By.id("treestop5")).click();
				driver.findElement(By.xpath("//*[@id='r5']/td[4]/a/img")).click();
				driver.switchTo().parentFrame();

				WebElement iframe3 = driver.findElement(By.xpath("/html/frameset/frameset[2]/frameset[2]/frame[2]"));
				driver.switchTo().frame(iframe3);
				driver.findElement(By.xpath("//*[@id='r1']/td[3]/a/img")).click();
				driver.switchTo().parentFrame();

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
				driver.findElement(By.id("fieldName:CUSTOMER:1")).sendKeys(documento.get(i));
				driver.findElement(By.id("fieldName:CURRENCY")).sendKeys("PEN");

				String screenshotPath1 = getScreenShot(driver, "");

				driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();

				wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:PRIMARY.OFFICER")));
				driver.findElement(By.id("fieldName:PRIMARY.OFFICER")).sendKeys(ejecutivo.get(i));
				driver.findElement(By.xpath("/html/body/div[5]/fieldset[3]/div/div/form[1]/div[3]/table/tbody/tr[2]/td/table/tbody/tr[8]/td[3]/table/tbody/tr/td[2]/input")).click();
				driver.findElement(By.xpath("/html/body/div[5]/fieldset[3]/div/div/form[1]/div[3]/table/tbody/tr[2]/td/table/tbody/tr[12]/td[3]/table/tbody/tr/td[1]/input")).click();


				String cod = driver.findElement(By.id("disabled_ACCOUNT.REFERENCE")).getText();
				System.out.println("CUENTA : " +cod);
				String arreglo = driver.findElement(By.id("disabled_ARRANGEMENT")).getText();
				System.out.println("ARREGLO : " +arreglo);
				write(i+1, 10, cod);

				wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:AMOUNT")));
				driver.findElement(By.id("fieldName:AMOUNT")).sendKeys(monto.get(i));

				String screenshotPath2 = getScreenShot(driver, "");

				driver.findElement(By.id("fieldName:CHANGE.PERIOD")).sendKeys(plazo.get(i));

				String screenshotPath3 = getScreenShot(driver, "");

				Select selectProducto = new Select(driver.findElement(By.id("fieldName:PAYIN.SETTLEMENT:1")));
				selectProducto.selectByVisibleText("YES");
				Select selectProducto2 = new Select(driver.findElement(By.id("fieldName:PAYOUT.SETTLEMENT:1")));
				selectProducto2.selectByVisibleText("YES");

				driver.findElement(By.id("fieldName:PAYIN.ACCOUNT:1:1")).sendKeys(cuenta.get(i));
				driver.findElement(By.id("fieldName:PAYOUT.ACCOUNT:1:1")).sendKeys(cuenta.get(i));

				String screenshotPath4 = getScreenShot(driver, "");

				Thread.sleep(1000);

				driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();


				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']")));
				driver.findElement(By.xpath("//img[@alt='Commit the deal']")).click();

				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div[3]/div/div[2]/form[1]/div[3]/table/tbody/tr[2]/td/table/tbody/tr/td[3]/select")));

				Select selectProducto1 = new Select(driver.findElement(By.xpath("/html/body/div[3]/div/div[2]/form[1]/div[3]/table/tbody/tr[2]/td/table/tbody/tr/td[3]/select")));
				selectProducto1.selectByVisibleText("RECEIVED");
				driver.findElement(By.id("errorImg")).click();

				//String cod = driver.findElement(By.id("transactionId")).getCssValue("value");
				String cod1 = driver.findElement(By.xpath("//*[@id='messages']/tbody/tr[2]/td[2]/table[2]/tbody/tr/td")).getText();
				String sSubCadena = cod1.substring(22,39);
				System.out.println(sSubCadena);
				write(i+1, 8, sSubCadena);

				String screenshotPath5 = getScreenShot(driver, "");


				logger.log(Status.PASS, MarkupHelper.createLabel("Dni Agregado", ExtentColor.GREEN));
				logger.log(Status.PASS,"Dni Agregado", MediaEntityBuilder.createScreenCaptureFromPath(screenshotPath1).build());
				logger.log(Status.PASS, MarkupHelper.createLabel("Monto Agregado", ExtentColor.GREEN));
				logger.log(Status.PASS,"Monto Agregado", MediaEntityBuilder.createScreenCaptureFromPath(screenshotPath2).build());
				logger.log(Status.PASS, MarkupHelper.createLabel("Plazo Agregado", ExtentColor.GREEN));
				logger.log(Status.PASS,"Plazo Agregado", MediaEntityBuilder.createScreenCaptureFromPath(screenshotPath3).build());
				logger.log(Status.PASS, MarkupHelper.createLabel("Cuenta Agregada", ExtentColor.GREEN));
				logger.log(Status.PASS,"Cuenta Agregada", MediaEntityBuilder.createScreenCaptureFromPath(screenshotPath4).build());
				logger.log(Status.PASS, MarkupHelper.createLabel("Alta creada", ExtentColor.GREEN));
				logger.log(Status.PASS,"Alta creada", MediaEntityBuilder.createScreenCaptureFromPath(screenshotPath5).build());
				extent.flush();
				write(i+1, 7, "PASSED");

				DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
				String fecha = dateFormat.format(new Date());
				System.out.println(fecha);
				write(i+1, 9, fecha);
					driver.quit();

				}

			}catch (Exception e){

				  String screenshotPath = getScreenShot(driver, "Error");
				  logger.log(Status.FAIL, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Error: "+e, ExtentColor.RED));
				  extent.flush();
				  write(i+1, 7, "FAILED");
				  write(i+1, 8, "");

				  DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
				  String fecha = dateFormat.format(new Date());
				  System.out.println(fecha);
				  write(i+1, 9, fecha);
				  System.out.println("Error: " + e);
				driver.quit();

			}
  		}

}

	public static ArrayList<String> readExcelData(int colNo) throws IOException {

		FileInputStream fis=new FileInputStream(System.getProperty("user.dir") + "/src/Excel/entregable1/AltaDPF.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet s=wb.getSheet("AltaDPF");
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
		String path = System.getProperty("user.dir") + "/src/Excel/entregable1/AltaDPF.xlsx";
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
		String destination = System.getProperty("user.dir") + "/test-output/reports/AltaDPF/Images/" + screenshotName + dateName + ".png";
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
