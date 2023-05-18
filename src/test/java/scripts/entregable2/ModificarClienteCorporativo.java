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
import java.time.Duration;


public class ModificarClienteCorporativo {

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
	public void ModificarClienteCorporativo()throws IOException, InterruptedException, AWTException {

		extent = new ExtentReports();
		spark = new ExtentSparkReporter(System.getProperty("user.dir") + "/test-output/reports2/ModificarClienteCorporativo/Report.html");
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
		ArrayList<String> RUC =readExcelData(2);
		ArrayList<String> RUCmod =readExcelData(3);

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
					driver.switchTo().parentFrame();

					// When El usuario da click en Menu
					Thread.sleep(1000);
					WebElement iframe2 = driver.findElement(By.xpath("/html/frameset/frame[2]"));
					driver.switchTo().frame(iframe2);
					driver.findElement(By.id("imgError")).click();

					//  And El usuario da click en Cliente
					driver.findElement(By.xpath("//img[@alt='Cliente']")).click();

					// El usuario da click en Modificar cliente
					driver.findElement(By.xpath("/html/body/div[3]/ul/li/ul/li[4]/ul/li[7]/a")).click();

					String MainWindow = driver.getWindowHandle();
					Set<String> s1 = driver.getWindowHandles();
					Iterator<String> i1 = s1.iterator();

					while (i1.hasNext()) {
						String ChildWindow = i1.next();

						if (!MainWindow.equalsIgnoreCase(ChildWindow)) {
							driver.switchTo().window(ChildWindow);
						}
					}

					//And  El usuario ingresa numero de documento RUC
					String attr1 = driver.findElement(By.xpath("//label[contains(text(),'Número de documento')]")).getAttribute("for");
					driver.findElement(By.id("value:1:1:1")).clear();
					driver.findElement(By.id("value:2:1:1")).clear();
					driver.findElement(By.id(attr1)).sendKeys(RUC.get(i));

					Thread.sleep(2000);

					//And El usuario da clikc en Find
					driver.findElement(By.xpath("//a[@alt='Run Selection']")).click();
					//And El usuario ingresa a editar cliente
					driver.findElement(By.xpath("/html/body/div[3]/div/form/div/table/tbody/tr[2]/td[2]/div[2]/div/table[1]/tbody/tr/td[6]/a/img")).click();

					Thread.sleep(5000);

					//And El usuario modifica al cliente y RUC

					driver.findElement(By.id("fieldName:SHORT.NAME:1")).clear();
					driver.findElement(By.id("fieldName:SHORT.NAME:1")).sendKeys("CopitoSAC");

					driver.findElement(By.id("fieldName:LEGAL.ID:1")).clear();
					driver.findElement(By.id("fieldName:LEGAL.ID:1")).sendKeys(RUCmod.get(i));


					//And El usuario prevalida
					driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();

					Thread.sleep(60000);

					//And El usuario guarda la modificacion
					driver.findElement(By.xpath("/html/body/div[3]/div[2]/form[1]/div[2]/table[2]/thead/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/div/table/tbody/tr/td[1]/a/img")).click();

					String cod = driver.findElement(By.xpath("//*[@id=\"messages\"]/tbody/tr[2]/td[2]/table[2]/tbody/tr/td")).getText();
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

	public void write(int i, int celda, String dato) throws IOException {
		String path = System.getProperty("user.dir") + "/src/Excel/entregable2/ModificarClienteCorporativo.xlsx";
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

	public void writeArreglo(int i, int celda, String dato) throws IOException {
		String path = System.getProperty("user.dir") + "/src/Excel/entregable2/ModificarClienteCorporativo.xlsx";
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
		String destination = System.getProperty("user.dir") + "/test-output/reports2/ModificarClienteCorporativo/Images/" + screenshotName + dateName + ".png";
		File finalDestination = new File(destination);
		FileUtils.copyFile(source, finalDestination);
		return destination;
	}

	public static ArrayList<String> readExcelData(int colNo) throws IOException {
		
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir") + "/src/Excel/entregable2/ModificarClienteCorporativo.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet s=wb.getSheet("ModificarClienteCorporativo");
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
