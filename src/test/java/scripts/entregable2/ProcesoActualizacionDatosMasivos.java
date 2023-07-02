package scripts.entregable2;

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


public class ProcesoActualizacionDatosMasivos {

    public static WebDriver driver;
	public ExtentSparkReporter spark;
	public ExtentReports extent;
	public ExtentTest logger;
	@BeforeTest
    public void startTest() {

    }
        
	@BeforeMethod
    public void openApplication() {
    	//System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir") + "/src/main/resources/drivers/chromedriver");
		System.setProperty("webdriver.gecko.driver", System.getProperty("user.dir") + "/src/main/resources/drivers/firefox");

	}


	@Test
	public void ProcesoActualizacionDatosMasivos()throws IOException, InterruptedException, AWTException {

		extent = new ExtentReports();
		spark = new ExtentSparkReporter(System.getProperty("user.dir") + "/test-output/reports2/ProcesoActualizacionDatosMasivos/Report.html");
		extent.attachReporter(spark);
		extent.setSystemInfo("Host Name", "SoftwareTestingMaterial");
		extent.setSystemInfo("Environment", "Production");
		extent.setSystemInfo("User Name", "Rajkumar SM");
		spark.config().setDocumentTitle("Title of the Report Comes here ");
		spark.config().setReportName("Name of the Report Comes here ");
		spark.config().setTheme(Theme.STANDARD);


    	Thread.sleep(1500);

		ArrayList<String> usuario= readExcelData(0);
		ArrayList<String> contraseña =readExcelData(1);
		ArrayList<String> archivo =readExcelData(2);
		ArrayList<String> frecuencia =readExcelData(3);


		int filas=usuario.size();
  		for(int i=0;i<usuario.size();i++) {
			try {
  			if(i<(filas)) {


					System.out.println("-----------------------------------");
					System.out.println("Nuevo Test " + i);
					int caso = i+1;
					logger = extent.createTest("Nuevo Test " + caso);

					// ** DESDE AQUI EMPIEZA EL TEST

				//  Given El usuario ingresa al Login Page
					driver = new ChromeDriver();

					//driver = new FirefoxDriver();

					driver.manage().window().maximize();
					driver.get("https://10.167.21.100:8480/BrowserWebSAD/servlet/BrowserServlet?");

					Thread.sleep(1000);
					driver.findElement(By.id("details-button")).click();
					driver.findElement(By.id("proceed-link")).click();

					// When El usuario ingresa el "<usuario>" y "<contraseña>"
					WebDriverWait wait = new WebDriverWait(driver, 40);
					wait.until(ExpectedConditions.elementToBeClickable(By.id("signOnName")));
					driver.findElement(By.id("signOnName")).sendKeys(usuario.get(i));
					driver.findElement(By.id("password")).sendKeys(contraseña.get(i));
					driver.findElement(By.id("sign-in")).click();

					WebElement iframe = driver.findElement(By.xpath("/html/frameset/frame[1]"));
					driver.switchTo().frame(iframe);

					// Then Redirecciona al Home Page
					Thread.sleep(2000);
					String exp_message = "Sign Off";
					String actual = driver.findElement(By.xpath("//a[contains(text(),'Sign Off')]")).getText();
					Assert.assertEquals(exp_message, actual);
					System.out.println("assert complete");


					// When El usuario ingresa codigo menu
					driver.findElement(By.id("commandValue")).sendKeys("?BRIP.RQ.MASS.UPD.CUS");
					driver.findElement(By.id("cmdline_img")).click();
					driver.switchTo().parentFrame();


					//El usuario Ingresa a la actualización
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

					//El usuario Ingresa a Cargar Archivo
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Carga de Archivos Actualizacion ')]"))).click();

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
					wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:DESCRIPTION"))).sendKeys("Descripcion 1");
					WebElement iframe1 = driver.findElement(By.xpath("/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[3]/td/table[1]/tbody/tr[3]/td[3]/iframe"));
					driver.switchTo().frame(iframe1);
					driver.findElement(By.xpath("//input[@type='file']")).sendKeys(System.getProperty("user.dir") + "/src/Excel/entregable2/Mod_Masiva_Clientes_PN_18.csv");
					driver.findElement(By.xpath("//img[@title='Upload']")).click();
					Thread.sleep(3000);
					driver.switchTo().parentFrame();


					driver.findElement(By.xpath("//img[@alt='Commit the deal']")).click();
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='messages']/tbody/tr[2]/td[2]/table[2]/tbody/tr/td")));
					driver.close();
					driver.switchTo().window(MainWindow2);

					//El usuario Ingresa a Servicio actualizacion
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Servicio de Actualizacion ')]"))).click();


					//El usuario Ingresa a la detalle actualización
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

					wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:USER")));
					wait.until(ExpectedConditions.elementToBeClickable(By.id("radio:mainTab:SERVICE.CONTROL"))).click();
					wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:FREQUENCY"))).clear();
					wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:USER"))).clear();
					wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:USER"))).sendKeys(frecuencia.get(i));

					//El usuario hace commit
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']")));
					driver.findElement(By.xpath("//img[@alt='Commit the deal']")).click();
					Thread.sleep(5000);


					String screenshotPath = getScreenShot(driver, "Fin del Caso");
					logger.log(Status.PASS, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Fin del Caso", ExtentColor.GREEN));
					extent.flush();
					write(i+1, 4, "PASSED");

					DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
					String fecha = dateFormat.format(new Date());
					System.out.println(fecha);
					write(i+1, 5, fecha);

					driver.quit();
				}

			}catch (Exception e){
				String screenshotPath = getScreenShot(driver, "Error");
				logger.log(Status.FAIL, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Error: "+e, ExtentColor.RED));
				extent.flush();
				write(i+1, 4, "FAILED");

				DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
				String fecha = dateFormat.format(new Date());
				System.out.println(fecha);
				write(i+1, 5, fecha);
				System.out.println("Error: " + e);
				driver.quit();
			}
  		}

}

	public void write(int i, int celda, String dato) throws IOException {
		String path = System.getProperty("user.dir") + "/src/Excel/entregable2/ProcesoActualizacionDatosMasivos.xlsx";
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
		String destination = System.getProperty("user.dir") + "/test-output/reports2/ProcesoActualizacionDatosMasivos/Images/" + screenshotName + dateName + ".png";
		File finalDestination = new File(destination);
		FileUtils.copyFile(source, finalDestination);
		return destination;
	}

	public static ArrayList<String> readExcelData(int colNo) throws IOException {
		
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir") + "/src/Excel/entregable2/ProcesoActualizacionDatosMasivos.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet s=wb.getSheet("ProcesoActualizacionDatosMasivo");
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
