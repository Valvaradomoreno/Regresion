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

public class PagoCuotaPrestamoCargoCuenta {


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
    public void PagoCuotaPrestamoCargoCuenta()throws IOException, InterruptedException, AWTException {

        extent = new ExtentReports();
        spark = new ExtentSparkReporter(System.getProperty("user.dir") + "/test-output/reports3/PagoCuotaPrestamoCargoCuenta/Report.html");
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
        ArrayList<String> cuenta =readExcelData(2);
        ArrayList<String> cuentaDeb =readExcelData(3);
        ArrayList<String> moneda =readExcelData(4);
        ArrayList<String> usuario2= readExcelData(5);


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

                    //El usuario da click Pago de cuota de prestamo
                    Thread.sleep(1000);
                    driver.findElement(By.xpath("//span[contains(text(),'Operaciones Minoristas')]")).click();
                    driver.findElement(By.xpath("//span[contains(text(),'Orden de Pago')]")).click();
                    driver.findElement(By.xpath("//span[contains(text(),'Input Payment Order')]")).click();
                    driver.findElement(By.xpath("//a[contains(text(),'Pago prestamo con cargo a cuenta ')]")).click();
                    driver.switchTo().parentFrame();

                    //El usuario busca la cuenta
                    String MainWindow = driver.getWindowHandle();
                    Set<String> s1 = driver.getWindowHandles();
                    Iterator<String> i1 = s1.iterator();

                    while (i1.hasNext()) {
                        String ChildWindow = i1.next();

                        if (!MainWindow.equalsIgnoreCase(ChildWindow)) {
                            driver.switchTo().window(ChildWindow);
                        }
                    }
                    wait.until(ExpectedConditions.elementToBeClickable(By.id("value:1:1:1")));
                    String attr = driver.findElement(By.xpath("//label[contains(text(),'Número de cuenta')]")).getAttribute("for");
                    driver.findElement(By.id("value:1:1:1")).clear();
                    Thread.sleep(200);
                    driver.findElement(By.id(attr)).sendKeys(cuenta.get(i));
                    driver.findElement(By.xpath("//a[@alt='Run Selection']")).click();

                    //El hace click en pago de cuota
                    driver.manage().window().maximize();
                    Thread.sleep(2000);
                    driver.findElement(By.xpath("//a[@alt='Select Drilldown']")).click();
                    Thread.sleep(2000);

                    //El usuario digita datos para pagar
                    driver.findElement(By.id("fieldName:DEBIT.ACCOUNT")).sendKeys(cuentaDeb.get(i));
                    driver.findElement(By.id("fieldName:DEBIT.CCY")).sendKeys(moneda.get(i));
                    driver.findElement(By.id("fieldName:PAYMENT.CURRENCY")).sendKeys(moneda.get(i));

                    //El usuario prevalida
                    driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();

                    //El usuario hace commit
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']")));
                    driver.findElement(By.xpath("//img[@alt='Commit the deal']")).click();

                    //Se muestra el codigo de transacción de alta prestamo
                    String cod2 = driver.findElement(By.xpath("//*[@id='messages']/tbody/tr[2]/td[2]/table[2]/tbody/tr/td")).getText();
                    String sSubCadena = cod2.substring(22,39);
                    System.out.println(sSubCadena);
                    write(i+1, 7, sSubCadena);


                    ////// APROBACION ******************


                    driver.get("https://10.167.21.100:8480/BrowserWebSAD/servlet/BrowserServlet?");

                    Thread.sleep(1000);
                    driver.findElement(By.id("details-button")).click();
                    driver.findElement(By.id("proceed-link")).click();
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

                    driver.findElement(By.xpath("//span[contains(text(),'Operaciones Minoristas')]")).click();
                    driver.findElement(By.xpath("//span[contains(text(),'Orden de Pago')]")).click();
                    driver.findElement(By.xpath("//span[contains(text(),'Input Payment Order')]")).click();
                    driver.findElement(By.xpath("//a[contains(text(),'Pago prestamo con cargo a cuenta ')]")).click();
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
                    String attr1 = driver.findElement(By.xpath("//label[contains(text(),'Número de cuenta')]")).getAttribute("for");
                    driver.findElement(By.id(attr1)).sendKeys(cuenta.get(i));
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

                    driver.findElement(By.xpath("//img[@alt='Select Drilldown']")).click();

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

                    String screenshotPath = getScreenShot(driver, "Fin del Caso");
                    logger.log(Status.PASS, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Fin del Caso", ExtentColor.GREEN));
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

    public void write(int i, int celda, String dato) throws IOException {
        String path = System.getProperty("user.dir") + "/src/Excel/entregable3/PagoCuotaPrestamoCargoCuenta.xlsx";
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
        String destination = System.getProperty("user.dir") + "/test-output/reports3/PagoCuotaPrestamoCargoCuenta/Images/" + screenshotName + dateName + ".png";
        File finalDestination = new File(destination);
        FileUtils.copyFile(source, finalDestination);
        return destination;
    }

    public static ArrayList<String> readExcelData(int colNo) throws IOException {

        FileInputStream fis=new FileInputStream(System.getProperty("user.dir") + "/src/Excel/entregable3/PagoCuotaPrestamoCargoCuenta.xlsx");
        XSSFWorkbook wb=new XSSFWorkbook(fis);
        XSSFSheet s=wb.getSheet("PagoCuotaPrestamoCargoCuenta");
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

