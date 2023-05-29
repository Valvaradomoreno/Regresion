package scripts.entregable3;

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

public class CuentaPagarDolares {

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
    public void CuentaPagarDolares()throws IOException, InterruptedException, AWTException {

        extent = new ExtentReports();
        spark = new ExtentSparkReporter(System.getProperty("user.dir") + "/test-output/reports3/CuentaPagarDolares/Report.html");
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

                    //El usuario da click en Operaciones Minoristas
                    Thread.sleep(1000);
                    driver.findElement(By.xpath("//span[contains(text(),'Operaciones Minoristas')]")).click();

                    //El usuario da click en Transacciones de Cuenta
                    Thread.sleep(500);
                    driver.findElement(By.xpath("//span[contains(text(),'Transacciones de Cuenta')]")).click();

                    //El usuario da click en Cajero
                    Thread.sleep(500);
                    driver.findElement(By.xpath("//span[contains(text(),'Cajero')]")).click();

                    //El usuario da click en Operaciones de Cajero
                    Thread.sleep(500);
                    driver.findElement(By.xpath("//img[@alt='Operaciones de Cajero']")).click();

                    //El usuario da click en Efectivo de Cajero
                    Thread.sleep(500);
                    driver.findElement(By.xpath("//span[contains(text(),'Efectivo de Cajero')]")).click();

                    //El usuario da click en RET.CTAXPAGAR.SOLES
                    Thread.sleep(500);
                    driver.findElement(By.xpath("//a[contains(text(),'RET.CTAXPAGAR.DOLARES ')]")).click();

                    driver.switchTo().parentFrame();

                    //El usuario ingresa a consumo banco ripley
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

                   driver.findElement(By.id("dealtitle")).isDisplayed();
                    Thread.sleep(1000);


                    String screenshotPath = getScreenShot(driver, "Fin del Caso");
                    logger.log(Status.PASS, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Fin del Caso", ExtentColor.GREEN));
                    extent.flush();
                    write(i+1, 2, "PASSED");

                    DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
                    String fecha = dateFormat.format(new Date());
                    System.out.println(fecha);
                    write(i+1, 3, fecha);

                    driver.quit();
                }

            }catch (Exception e){
                String screenshotPath = getScreenShot(driver, "Error");
                logger.log(Status.FAIL, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Error: "+e, ExtentColor.RED));
                extent.flush();
                write(i+1, 2, "FAILED");

                DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
                String fecha = dateFormat.format(new Date());
                System.out.println(fecha);
                write(i+1, 3, fecha);
                System.out.println("Error: " + e);
                driver.quit();
            }
        }

    }

    public void write(int i, int celda, String dato) throws IOException {
        String path = System.getProperty("user.dir") + "/src/Excel/entregable3/CuentaPagarDolares.xlsx";
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
        String destination = System.getProperty("user.dir") + "/test-output/reports3/CuentaPagarDolares/Images/" + screenshotName + dateName + ".png";
        File finalDestination = new File(destination);
        FileUtils.copyFile(source, finalDestination);
        return destination;
    }

    public static ArrayList<String> readExcelData(int colNo) throws IOException {

        FileInputStream fis=new FileInputStream(System.getProperty("user.dir") + "/src/Excel/entregable3/CuentaPagarDolares.xlsx");
        XSSFWorkbook wb=new XSSFWorkbook(fis);
        XSSFSheet s=wb.getSheet("CuentaPagarDolares");
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
