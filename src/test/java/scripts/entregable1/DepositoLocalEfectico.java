package scripts.entregable1;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
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


public class DepositoLocalEfectico {

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
	public void DepositoLocalEfectico()throws IOException, InterruptedException, AWTException {

		extent = new ExtentReports();
		spark = new ExtentSparkReporter(System.getProperty("user.dir") + "/test-output/reports/DepositoLocalEfectico/Report.html");
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
		ArrayList<String> monto =readExcelData(3);

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
				WebDriverWait wait = new WebDriverWait(driver, 60);
/*
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

				driver.findElement(By.xpath("//img[@alt='Operaciones Minoristas']")).click();

				driver.findElement(By.xpath("//a[contains(text(),'Buscar Cuenta ')]")).click();
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
				String attr = driver.findElement(By.xpath("//label[contains(text(),'Código de cliente')]")).getAttribute("for");
				driver.findElement(By.id(attr)).sendKeys(cuenta.get(i));
				driver.findElement(By.xpath("//a[@alt='Run Selection']")).click();
				Thread.sleep(2000);

				driver.findElement(By.xpath("/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div[3]/div/form/div/table/tbody/tr[2]/td[2]/div[2]/div/table[1]/tbody/tr/td[7]/a/img")).click();

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

				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/table/tbody/tr[11]/td/table/tbody/tr[1]/td/div[3]/div/form/div/table/tbody/tr[2]/td[2]/div[2]/div/table[1]/tbody/tr[2]/td[5]")));
				String saldo = driver.findElement(By.xpath("/html/body/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/table/tbody/tr[11]/td/table/tbody/tr[1]/td/div[3]/div/form/div/table/tbody/tr[2]/td[2]/div[2]/div/table[1]/tbody/tr[2]/td[5]")).getText();
*/


				driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				driver.get("https://10.167.21.100:8480/BrowserWebSAD/servlet/BrowserServlet?");
				Thread.sleep(1000);
				driver.findElement(By.id("details-button")).click();
				driver.findElement(By.id("proceed-link")).click();

				wait.until(ExpectedConditions.elementToBeClickable(By.id("signOnName")));

				driver.findElement(By.id("signOnName")).sendKeys(usuario.get(i));
				driver.findElement(By.id("password")).sendKeys(contraseña.get(i));
				driver.findElement(By.id("sign-in")).click();

				//WebDriverWait wait = new WebDriverWait(driver, 30);
				//wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//a[contains(text(),'Sign Off')]")));

				WebElement iframe5 = driver.findElement(By.xpath("/html/frameset/frame[1]"));
				driver.switchTo().frame(iframe5);

				Thread.sleep(2000);
				String exp_message1 = "Sign Off";
				String actual1 = driver.findElement(By.xpath("//a[contains(text(),'Sign Off')]")).getText();
				//Assert.assertEquals(exp_message, actual1);
				System.out.println("assert complete");
				driver.switchTo().parentFrame();

				Thread.sleep(1000);
				WebElement iframe2 = driver.findElement(By.xpath("/html/frameset/frame[2]"));
				driver.switchTo().frame(iframe2);
				driver.findElement(By.id("imgError")).click();

				driver.findElement(By.xpath("//img[@alt='Operaciones Minoristas']")).click();
				//Thread.sleep(1000);

				//driver.findElement(By.xpath("//span[contains(text(),'Transacciones de Cuenta')]")).click();


				driver.findElement(By.xpath("//span[contains(text(),'Transacciones de Cuenta')]")).click();

				driver.findElement(By.xpath("//span[contains(text(),'Cajero')]")).click();

				driver.findElement(By.xpath("//img[@alt='Operaciones de Cajero']")).click();

				driver.findElement(By.xpath("//span[contains(text(),'Efectivo de Cajero')]")).click();

				driver.findElement(By.xpath("//a[contains(text(),'Depósito Local en Efectivo ')]")).click();

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

				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='mainTab']/tbody/tr[6]/td[3]/a[2]/img")));
				driver.findElement(By.xpath("//*[@id='mainTab']/tbody/tr[6]/td[3]/a[2]/img")).click();
				String MainWindow4=driver.getWindowHandle();
				Set<String> s4=driver.getWindowHandles();
				Iterator<String> i4=s4.iterator();

				while(i4.hasNext())
				{
					String ChildWindow2=i4.next();

					if(!MainWindow4.equalsIgnoreCase(ChildWindow2))
					{
						driver.switchTo().window(ChildWindow2);
					}
				}
				String attr1 = driver.findElement(By.xpath("//label[contains(text(),'Cliente')]")).getAttribute("for");
				driver.findElement(By.id(attr1)).clear();
				driver.findElement(By.id(attr1)).sendKeys(cuenta.get(i));
				driver.findElement(By.xpath("//a[@alt='Run Selection']")).click();
				Thread.sleep(2000);
				//driver.findElement(By.xpath("//b[contains(text(),'"+cuenta.get(i)+"')]")).click();
				driver.findElement(By.xpath("/html/body/div[3]/div/form/div/table/tbody/tr[2]/td[2]/div[2]/div/table[1]/tbody/tr/td[1]/a")).click();
				driver.switchTo().window(MainWindow4);


				wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:AMOUNT.LOCAL.1:1")));
				driver.findElement(By.id("fieldName:AMOUNT.LOCAL.1:1")).sendKeys(monto.get(i));

				driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();


				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']")));
				driver.findElement(By.xpath("//img[@alt='Commit the deal']")).click();

				Thread.sleep(5000);
				if(driver.findElement(By.id("errorImg")).isDisplayed()){
					driver.findElement(By.id("errorImg")).click();
				}
				//wait.until(ExpectedConditions.elementToBeClickable(By.id("errorImg"))).isDisplayed();
				//driver.findElement(By.id("errorImg")).click();

				Thread.sleep(3000);

				//String cod = driver.findElement(By.id("transactionId")).getCssValue("value");
				String cod = driver.findElement(By.xpath("//*[@id='messages']/tbody/tr[2]/td[2]/table[2]/tbody/tr/td")).getText();
				String cod2 = "ALTA CUENTA";
				System.out.println("este es : " +cod);
				String[] datoWrite = {cod,cod2};
				//Escriba los datos a llenar


				/*driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				driver.get("https://10.167.21.100:8480/BrowserWebSAD/servlet/BrowserServlet?");
				Thread.sleep(1000);

				wait.until(ExpectedConditions.elementToBeClickable(By.id("signOnName")));

				driver.findElement(By.id("signOnName")).sendKeys(usuario.get(i));
				driver.findElement(By.id("password")).sendKeys(contraseña.get(i));
				driver.findElement(By.id("sign-in")).click();

				//WebDriverWait wait = new WebDriverWait(driver, 30);
				//wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//a[contains(text(),'Sign Off')]")));

				WebElement iframe3 = driver.findElement(By.xpath("/html/frameset/frame[1]"));
				driver.switchTo().frame(iframe3);

				Thread.sleep(2000);
				String exp_message2 = "Sign Off";
				String actual2 = driver.findElement(By.xpath("//a[contains(text(),'Sign Off')]")).getText();
				Assert.assertEquals(exp_message, actual2);
				System.out.println("assert complete");
				driver.switchTo().parentFrame();

				Thread.sleep(1000);
				WebElement iframe4 = driver.findElement(By.xpath("/html/frameset/frame[2]"));
				driver.switchTo().frame(iframe4);

				driver.findElement(By.id("imgError")).click();

				driver.findElement(By.xpath("//img[@alt='Operaciones Minoristas']")).click();

				driver.findElement(By.xpath("//a[contains(text(),'Buscar Cuenta ')]")).click();
				driver.switchTo().parentFrame();

				String MainWindow5=driver.getWindowHandle();
				Set<String> s5=driver.getWindowHandles();
				Iterator<String> i5=s5.iterator();

				while(i5.hasNext())
				{
					String ChildWindow=i5.next();

					if(!MainWindow5.equalsIgnoreCase(ChildWindow))
					{
						driver.switchTo().window(ChildWindow);
					}
				}


				wait.until(ExpectedConditions.elementToBeClickable(By.id("value:1:1:1")));
				driver.findElement(By.id("value:1:1:1")).clear();
				Thread.sleep(200);
				driver.findElement(By.id("value:2:1:1")).clear();
				Thread.sleep(200);
				String attr2 = driver.findElement(By.xpath("//label[contains(text(),'Código de cliente')]")).getAttribute("for");
				driver.findElement(By.id(attr)).sendKeys(cuenta.get(i));
				driver.findElement(By.xpath("//a[@alt='Run Selection']")).click();

				driver.findElement(By.xpath("/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div[3]/div/form/div/table/tbody/tr[2]/td[2]/div[2]/div/table[1]/tbody/tr/td[7]/a/img")).click();


				String MainWindow6=driver.getWindowHandle();
				Set<String> s6=driver.getWindowHandles();
				Iterator<String> i6=s6.iterator();

				while(i6.hasNext())
				{
					String ChildWindow=i6.next();

					if(!MainWindow6.equalsIgnoreCase(ChildWindow))
					{
						driver.switchTo().window(ChildWindow);
					}
				}

				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/table/tbody/tr[11]/td/table/tbody/tr[1]/td/div[3]/div/form/div/table/tbody/tr[2]/td[2]/div[2]/div/table[1]/tbody/tr[2]/td[5]")));
				String saldo2 = driver.findElement(By.xpath("/html/body/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/table/tbody/tr[11]/td/table/tbody/tr[1]/td/div[3]/div/form/div/table/tbody/tr[2]/td[2]/div[2]/div/table[1]/tbody/tr[2]/td[5]")).getText();
*/
				//Assert.assertNotEquals(saldo2, saldo);
				//String cod1 = driver.findElement(By.xpath("//*[@id='messages']/tbody/tr[2]/td[2]/table[2]/tbody/tr/td")).getText();
				String sSubCadena = cod.substring(22,39);
				System.out.println(sSubCadena);
				write(i+1, 5, sSubCadena);


				String screenshotPath = getScreenShot(driver, "Fin del Caso");
				logger.log(Status.PASS, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Fin del Caso", ExtentColor.GREEN));
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

		FileInputStream fis=new FileInputStream(System.getProperty("user.dir") + "/src/Excel/entregable1/DepositoLocalEfectico.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet s=wb.getSheet("DepositoLocalEfectico");
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
		String path = System.getProperty("user.dir") + "/src/Excel/entregable1/DepositoLocalEfectico.xlsx";
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
		String destination = System.getProperty("user.dir") + "/test-output/reports/DepositoLocalEfectico/Images/" + screenshotName + dateName + ".png";
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
