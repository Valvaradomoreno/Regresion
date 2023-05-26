package scripts.entregable3;

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
import java.time.Duration;


public class CambioT24aCastigo {

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
	public void CambioT24aCastigo()throws IOException, InterruptedException, AWTException {

		extent = new ExtentReports();
		spark = new ExtentSparkReporter(System.getProperty("user.dir") + "/test-output/reports3/CambioT24aCastigo/Report.html");
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
		ArrayList<String> tipo_producto =readExcelData(1);
		ArrayList<String> arreglo =readExcelData(3);
		ArrayList<String> producto =readExcelData(4);
		ArrayList<String> usuario2=readExcelData(5);


		int filas=usuario.size();
  		for(int i=0;i<usuario.size();i++) {
			  try {

  			if(i<(filas)) {

				System.out.println("-----------------------------------");
				System.out.println("Nuevo Test " + i);
				int caso = i + 1;
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

				driver.findElement(By.xpath("//a[contains(text(),'Buscar Préstamo ')]")).click();
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
				String attr = driver.findElement(By.xpath("//label[contains(text(),'ID de Arreglo')]")).getAttribute("for");
				driver.findElement(By.id(attr)).sendKeys(arreglo.get(i));
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
				Thread.sleep(2500);
				String screenshotPath1 = getScreenShot(driver, "");

				driver.findElement(By.xpath("//a[contains(text(),'Nueva Actividad')]")).click();

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


				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div[3]/div/form/div/table/tbody/tr[2]/td[2]/div[3]/div/table[1]/tbody/tr[1]/td[3]/a/img")));

				System.out.println("GOO");
				String xp = driver.findElement(By.xpath("//tr/td[contains(text(),'Cambio de Condiciones')]")).getAttribute("id");
				System.out.println("Id:" + xp);
				Thread.sleep(1000);

				// TUERCA
				if(tipo_producto.get(i).equals("PRESTAMO YA")){
					driver.findElement(By.xpath("/html/body/div[3]/div/form/div/table/tbody/tr[2]/td[2]/div[3]/div/table[1]/tbody/tr[36]/td[3]/a/img")).click();
				}else if(tipo_producto.get(i).equals("PRESTAMO SENIOR")){
					driver.findElement(By.xpath("/html/body/div[3]/div/form/div/table/tbody/tr[2]/td[2]/div[3]/div/table[1]/tbody/tr[36]/td[3]/a/img")).click();
				}else{
					driver.findElement(By.xpath("/html/body/div[3]/div/form/div/table/tbody/tr[2]/td[2]/div[3]/div/table[1]/tbody/tr[36]/td[3]/a/img")).click();
				}

				driver.manage().window().maximize();
				Thread.sleep(1000);

				// ELEGIR PRODUCTO
				wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:PRODUCT")));
				driver.findElement(By.id("fieldName:PRODUCT")).sendKeys(producto.get(i));
				String screenshotPath2 = getScreenShot(driver, "");

				driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();


				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']")));
				driver.findElement(By.xpath("//img[@alt='Commit the deal']")).click();

				wait.until(ExpectedConditions.elementToBeClickable(By.id("errorImg")));

				driver.findElement(By.id("errorImg")).click();

				Thread.sleep(3000);


				String cod = driver.findElement(By.xpath("//*[@id='messages']/tbody/tr[2]/td[2]/table[2]/tbody/tr/td")).getText();
				String sSubCadena = cod.substring(22,39);
				System.out.println(sSubCadena);
				write(i+1, 7, sSubCadena);


				////// APROBACION ******************


				driver.get("https://10.167.21.100:8480/BrowserWebSAD/servlet/BrowserServlet?");

				Thread.sleep(3000);

				driver.findElement(By.id("signOnName")).sendKeys(usuario2.get(i));
				driver.findElement(By.id("password")).sendKeys(contraseña.get(i));
				driver.findElement(By.id("sign-in")).click();

				WebElement iframe0 = driver.findElement(By.xpath("/html/frameset/frame[1]"));
				driver.switchTo().frame(iframe0);
				Assert.assertEquals(exp_message, actual);
				System.out.println("assert complete");
				driver.switchTo().parentFrame();

				Thread.sleep(1000);
				WebElement iframe10 = driver.findElement(By.xpath("/html/frameset/frame[2]"));
				driver.switchTo().frame(iframe10);

				driver.findElement(By.id("imgError")).click();

				driver.findElement(By.xpath("//img[@alt='Operaciones Minoristas']")).click();

				driver.findElement(By.xpath("//a[contains(text(),'Buscar Préstamo ')]")).click();
				driver.switchTo().parentFrame();

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

				Thread.sleep(3000);
				driver.findElement(By.id("value:1:1:1")).clear();
				Thread.sleep(200);
				driver.findElement(By.id("value:2:1:1")).clear();
				Thread.sleep(200);
				String attr1 = driver.findElement(By.xpath("//label[contains(text(),'ID de Arreglo')]")).getAttribute("for");
				driver.findElement(By.id(attr1)).sendKeys(arreglo.get(i));
				driver.findElement(By.xpath("//a[@alt='Run Selection']")).click();
				Thread.sleep(500);

				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Overview']"))).click();

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
				driver.manage().window().maximize();
				Thread.sleep(5000);

				String screenshotPath3 = getScreenShot(driver, "");
				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Select Drilldown']"))).click();

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

				driver.findElement(By.xpath("//img[@alt='Authorises a deal']")).click();
				Thread.sleep(3000);


				String screenshotPath4 = getScreenShot(driver, "");
				logger.log(Status.PASS, MarkupHelper.createLabel("Producto antes", ExtentColor.GREEN));
				logger.log(Status.PASS,"Producto antes", MediaEntityBuilder.createScreenCaptureFromPath(screenshotPath1).build());
				logger.log(Status.PASS, MarkupHelper.createLabel("Producto Castigo", ExtentColor.GREEN));
				logger.log(Status.PASS,"Producto Castigo", MediaEntityBuilder.createScreenCaptureFromPath(screenshotPath2).build());
				logger.log(Status.PASS, MarkupHelper.createLabel("Por aprobar", ExtentColor.GREEN));
				logger.log(Status.PASS,"Por aprobar", MediaEntityBuilder.createScreenCaptureFromPath(screenshotPath3).build());
				logger.log(Status.PASS, MarkupHelper.createLabel("Fin del Caso", ExtentColor.GREEN));
				logger.log(Status.PASS,"Fin del Caso", MediaEntityBuilder.createScreenCaptureFromPath(screenshotPath4).build());

				extent.flush();
				write(i+1, 6, "PASSED");

				DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
				String fecha = dateFormat.format(new Date());
				System.out.println(fecha);
				write(i+1, 8, fecha);
					driver.quit();

				}

			}catch (Exception e){

				  String screenshotPath = getScreenShot(driver, "Error");
				  logger.log(Status.FAIL, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Error: "+e, ExtentColor.RED));
				  extent.flush();
				  write(i+1, 6, "FAILED");
				  write(i+1, 7, "");

				  DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
				  String fecha = dateFormat.format(new Date());
				  System.out.println(fecha);
				  write(i+1, 8, fecha);
				  System.out.println("Error: " + e);
				driver.quit();

			}
  		}

}

	public static ArrayList<String> readExcelData(int colNo) throws IOException {

		FileInputStream fis=new FileInputStream(System.getProperty("user.dir") + "/src/Excel/entregable3/CambioT24aCastigo.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet s=wb.getSheet("CambioT24aCastigo");
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
		String path = System.getProperty("user.dir") + "/src/Excel/entregable3/CambioT24aCastigo.xlsx";
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
		String destination = System.getProperty("user.dir") + "/test-output/reports3/CambioT24aCastigo/Images/" + screenshotName + dateName + ".png";
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
