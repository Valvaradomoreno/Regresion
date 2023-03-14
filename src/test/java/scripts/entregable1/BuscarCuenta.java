package scripts.entregable1;

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
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.AfterTest;
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
import java.util.*;


public class BuscarCuenta {

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
	public void buscarCuenta()throws IOException, InterruptedException, AWTException {

		extent = new ExtentReports();
		spark = new ExtentSparkReporter(System.getProperty("user.dir") + "/test-output/reports/BuscarCuenta/Report.html");
		extent.attachReporter(spark);
		extent.setSystemInfo("Host Name", "SoftwareTestingMaterial");
		extent.setSystemInfo("Environment", "Production");
		extent.setSystemInfo("User Name", "Rajkumar SM");
		spark.config().setDocumentTitle("Title of the Report Comes here ");
		spark.config().setReportName("Name of the Report Comes here ");
		spark.config().setTheme(Theme.STANDARD);

	ArrayList<String> usuario = readExcelData(0);
	ArrayList<String> contraseña = readExcelData(1);
	ArrayList<String> arreglo = readExcelData(2);

	int filas = usuario.size();
	for (int i = 0; i < usuario.size(); i++) {
		try {

		if (i < (filas)) {


				System.out.println("-----------------------------------");
				System.out.println("Nuevo Test " + i);
				int caso = i+1;
				logger = extent.createTest("Nuevo Test " + caso);

				// ** EMPIEZA TEST

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
				WebElement iframe2 = driver.findElement(By.xpath("/html/frameset/frame[2]"));
				driver.switchTo().frame(iframe2);
				driver.findElement(By.id("imgError")).click();

				driver.findElement(By.xpath("//img[@alt='Operaciones Minoristas']")).click();
				driver.findElement(By.xpath("//a[contains(text(),'Buscar Cuenta ')]")).click();
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


				System.out.println("es: " + arreglo);

				wait.until(ExpectedConditions.elementToBeClickable(By.id("value:1:1:1")));
				String attr = driver.findElement(By.xpath("//label[contains(text(),'Número de cuenta')]")).getAttribute("for");
				driver.findElement(By.id("value:1:1:1")).clear();
				Thread.sleep(200);
				driver.findElement(By.id("value:2:1:1")).clear();
				Thread.sleep(200);
				System.out.println("arreglo: " + arreglo.get(i));
				driver.findElement(By.id(attr)).sendKeys(arreglo.get(i));

			String screenshotPath = getScreenShot(driver, "Cuenta a buscar");


			driver.findElement(By.xpath("//a[@alt='Run Selection']")).click();

				driver.findElement(By.xpath("/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/div[3]/div/form/div/table/tbody/tr[2]/td[2]/div[2]/div/table[1]/tbody/tr[1]/td[7]/a/img")).click();
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
				wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//td[contains(text(),'Authorised')]")));
				Thread.sleep(6000);
				String screenshotPath1 = getScreenShot(driver, "Cuenta encontrada");
				System.out.println("Agregar Excel");
				write(i+1, 3, "PASSED");

				logger.log(Status.PASS, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Cuenta a buscar", ExtentColor.GREEN));
				logger.log(Status.PASS, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath1) + " Cuenta encontrada", ExtentColor.GREEN));

				extent.flush();

				System.out.println("rowCount: " + i);
				DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
				String date = dateFormat.format(new Date());
				write(i+1, 4, date);
				System.out.println(date);
				driver.quit();
			}

		}catch (Exception e){
			String screenshotPath = getScreenShot(driver, "ERROR");
			logger.log(Status.FAIL, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Test Case FAIL: "+e, ExtentColor.RED));
			//objExcelFile.writeToExcel("FAIL", rowCount,4);
			write(i+1, 3, "FAILED");
			extent.flush();

			System.out.println("rowCount: " + i);
			DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
			String date = dateFormat.format(new Date());
			write(i+1, 4, date);
			driver.quit();
		}
	}


}
	public void write(int i, int celda, String dato) throws IOException {
		String path = System.getProperty("user.dir") + "/src/Excel/BuscarCuenta.xlsx";
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
		String destination = System.getProperty("user.dir") + "/reports/BuscarCuenta/Images/" + screenshotName + dateName + ".png";
		File finalDestination = new File(destination);
		FileUtils.copyFile(source, finalDestination);
		return destination;
	}
	public static ArrayList<String> readExcelData(int colNo) throws IOException {
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir") + "/src/Excel/BuscarCuenta.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet s=wb.getSheet("Buscar Cuenta");
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
