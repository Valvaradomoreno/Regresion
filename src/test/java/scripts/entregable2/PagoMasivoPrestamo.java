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
import java.time.Duration;
import java.util.concurrent.TimeUnit;

public class PagoMasivoPrestamo {


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
    public void PagoMasivoPrestamo()throws IOException, InterruptedException, AWTException {

        extent = new ExtentReports();
        spark = new ExtentSparkReporter(System.getProperty("user.dir") + "/test-output/reports2/PagoMasivoPrestamo/Report.html");
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
        ArrayList<String> tipo_carga =readExcelData(2);
        ArrayList<String> archivo =readExcelData(3);
        ArrayList<String> usuarioAp= readExcelData(4);
        ArrayList<String> cuenta= readExcelData(5);



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

                    driver.findElement(By.xpath("//img[@alt='Servicios de Pagos']")).click();
                    driver.findElement(By.xpath("//img[@alt='Pagos Masivos']")).click();
                    driver.findElement(By.xpath("//img[@alt='Creación de FT Masivo Master']")).click();

                    driver.findElement(By.xpath("//a[contains(text(),'Carga de Archivo Masivo ')]")).click();
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

                    Thread.sleep(3000);
                    WebElement iframe3 = driver.findElement(By.xpath("/html/frameset/frame[2]"));
                    driver.switchTo().frame(iframe3);

                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='New Deal']"))).click();
                    wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:DESCRIPTION"))).sendKeys("Descripción 1");
                    wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:UPLOAD.TYPE"))).sendKeys(tipo_carga.get(i));

                    WebElement iframe4 = driver.findElement(By.xpath("/html/body/div[3]/div[2]/form[1]/div[4]/table/tbody/tr[3]/td/table[1]/tbody/tr[3]/td[3]/iframe"));
                    driver.switchTo().frame(iframe4);
                    Thread.sleep(2000);
                    driver.findElement(By.xpath("//input[@type='file']")).sendKeys(System.getProperty("user.dir") +archivo.get(i));
                    Thread.sleep(2000);
                    driver.findElement(By.xpath("//img[@title='Upload']")).click();
                    Thread.sleep(1000);
                    driver.switchTo().parentFrame();
                    driver.switchTo().parentFrame();

                    WebElement iframe5 = driver.findElement(By.xpath("/html/frameset/frame[2]"));
                    driver.switchTo().frame(iframe5);
                    driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']")));
                    driver.findElement(By.xpath("//img[@alt='Commit the deal']")).click();

                    Thread.sleep(3000);


                    // START *****
                    driver.switchTo().window(MainWindow);
                    driver.switchTo().frame(iframe2);
                    driver.findElement(By.xpath("//a[contains(text(),'Establecer Servicio TSA T24.UPLOAD.PROCESS ')]")).click();
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
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@value='START']"))).click();
                    //driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']")));
                    driver.findElement(By.xpath("//img[@alt='Commit the deal']")).click();
                    Thread.sleep(5000);

                    // STOP *****
                   /* driver.switchTo().window(MainWindow);
                    driver.switchTo().frame(iframe2);
                    driver.findElement(By.xpath("//a[contains(text(),'Establecer Servicio TSA T24.UPLOAD.PROCESS ')]")).click();
                    driver.switchTo().parentFrame();

                    String MainWindow21=driver.getWindowHandle();
                    Set<String> s21=driver.getWindowHandles();
                    Iterator<String> i21=s21.iterator();

                    while(i21.hasNext())
                    {
                        String ChildWindow=i21.next();

                        if(!MainWindow21.equalsIgnoreCase(ChildWindow))
                        {
                            driver.switchTo().window(ChildWindow);
                        }
                    }


                    driver.manage().window().maximize();
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[contains(text(),'STOP')]"))).click();
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']")));
                    driver.findElement(By.xpath("//img[@alt='Commit the deal']")).click();
                    Thread.sleep(3000);*/

                    // VALIDAR ****

                    driver.switchTo().window(MainWindow);
                    driver.switchTo().frame(iframe2 );
                    driver.findElement(By.xpath("//a[contains(text(),'Ingresar/Validar Masivo Maestro ')] ")).click();
                    driver.switchTo().parentFrame();

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

                    driver.manage().window().maximize();
                    Thread.sleep(1500);
                    WebElement iframe11 = driver.findElement(By.xpath("/html/frameset/frame[1]"));
                    driver.switchTo().frame(iframe11);
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@title='Selection Screen']"))).click();
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//label[contains(text(),'Cuenta Activa')]")));
                    String attr1 = driver.findElement(By.xpath("//label[contains(text(),'Cuenta Activa')]")).getAttribute("for");
                    driver.findElement(By.id(attr1)).clear();
                    driver.findElement(By.id(attr1)).sendKeys(cuenta.get(i));
                    driver.findElement(By.xpath("//a[@alt='Run Selection']")).click();
                    Thread.sleep(1000);
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@title='Validate']"))).click();
                    driver.switchTo().parentFrame();
                    Thread.sleep(2500);
                    WebElement iframe12 = driver.findElement(By.xpath("/html/frameset/frame[2]"));
                    driver.switchTo().frame(iframe12);
                    DateFormat dateFormat1 = new SimpleDateFormat("yyyyMMdd");
                    String fecha1 = dateFormat1.format(new Date());
                    System.out.println(fecha1);
                    wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:PROCESSING.DATE"))).clear();
                    wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:PROCESSING.DATE"))).sendKeys(fecha1);
                    wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:PAYMENT.VALUE.DATE"))).clear();
                    wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:PAYMENT.VALUE.DATE"))).sendKeys(fecha1);
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']"))).click();
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'Accept Overrides')]"))).click();
                    Thread.sleep(30000);
                    driver.switchTo().parentFrame();


                    ////// APROBACION ******************


                    driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
                    driver.get("https://10.167.21.100:8480/BrowserWebSAD/servlet/BrowserServlet?");
                    Thread.sleep(1000);

                    wait.until(ExpectedConditions.elementToBeClickable(By.id("signOnName")));

                    driver.findElement(By.id("signOnName")).sendKeys(usuarioAp.get(i));
                    driver.findElement(By.id("password")).sendKeys(contraseña.get(i));
                    driver.findElement(By.id("sign-in")).click();
                    driver.manage().window().maximize();

                    driver.switchTo().frame(iframe2);

                    Thread.sleep(2000);
                    String exp_message1 = "Sign Off";
                    String actual1 = driver.findElement(By.xpath("//a[contains(text(),'Sign Off')]")).getText();
                    Assert.assertEquals(exp_message, actual1);
                    System.out.println("assert complete");
                    driver.switchTo().parentFrame();

                    Thread.sleep(1000);
                    driver.switchTo().frame(iframe3);

                    driver.findElement(By.id("imgError")).click();

                    driver.findElement(By.xpath("//img[@alt='Servicios de Pagos']")).click();
                    driver.findElement(By.xpath("//img[@alt='Pagos Masivos']")).click();
                    driver.findElement(By.xpath("//img[@alt='Creación de FT Masivo Master']")).click();

                    driver.findElement(By.xpath("//a[contains(text(),'Autorizar Masivo Maestro ')]")).click();
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

                    driver.manage().window().maximize();
                    driver.switchTo().frame(iframe11);
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@title='Selection Screen']"))).click();
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//label[contains(text(),'Cuenta Activa')]")));
                    String attr2 = driver.findElement(By.xpath("//label[contains(text(),'Cuenta Activa')]")).getAttribute("for");
                    driver.findElement(By.id(attr2)).clear();
                    driver.findElement(By.id(attr2)).sendKeys(cuenta.get(i));
                    driver.findElement(By.xpath("//a[@alt='Run Selection']")).click();
                    Thread.sleep(1000);
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@title='Validate']"))).click();
                    driver.switchTo().parentFrame();
                    driver.switchTo().frame(iframe12);
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']")));
                    driver.findElement(By.xpath("//img[@alt='Commit the deal']")).click();
                    Thread.sleep(3000);
                    driver.switchTo().parentFrame();



                    // START *****

                    driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
                    driver.get("https://10.167.21.100:8480/BrowserWebSAD/servlet/BrowserServlet?");
                    driver.switchTo().window(MainWindow);
                    Thread.sleep(1000);


                    wait.until(ExpectedConditions.elementToBeClickable(By.id("signOnName")));

                    driver.findElement(By.id("signOnName")).sendKeys(usuarioAp.get(i));
                    driver.findElement(By.id("password")).sendKeys(contraseña.get(i));
                    driver.findElement(By.id("sign-in")).click();

                    driver.switchTo().frame(iframe2);

                    driver.findElement(By.id("imgError")).click();

                    driver.findElement(By.xpath("//img[@alt='Servicios de Pagos']")).click();
                    driver.findElement(By.xpath("//img[@alt='Pagos Masivos']")).click();

                    driver.findElement(By.xpath("//a[contains(text(),'Establecer Servicio TSA T24.UPLOAD.PROCESS ')]")).click();
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

                    driver.manage().window().maximize();
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@value='START']"))).click();
                    driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']")));
                    driver.findElement(By.xpath("//img[@alt='Commit the deal']")).click();
                    Thread.sleep(3000);


                    String cod = driver.findElement(By.xpath("//*[@id='messages']/tbody/tr[2]/td[2]/table[2]/tbody/tr/td")).getText();
                    String sSubCadena = cod.substring(22,39);
                    System.out.println(sSubCadena);
                    write(i+1, 7, sSubCadena);

                    extent.flush();
                    write(i+1, 6, "PASSED");


                    DateFormat dateFormat2 = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
                    String fecha = dateFormat2.format(new Date());
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
        String path = System.getProperty("user.dir") + "/src/Excel/entregable2/PagoMasivoPrestamo.xlsx";
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
        String destination = System.getProperty("user.dir") + "/test-output/reports2/PagoMasivoPrestamo/Images/" + screenshotName + dateName + ".png";
        File finalDestination = new File(destination);
        FileUtils.copyFile(source, finalDestination);
        return destination;
    }

    public static ArrayList<String> readExcelData(int colNo) throws IOException {

        FileInputStream fis=new FileInputStream(System.getProperty("user.dir") + "/src/Excel/entregable2/PagoMasivoPrestamo.xlsx");
        XSSFWorkbook wb=new XSSFWorkbook(fis);
        XSSFSheet s=wb.getSheet("PagoMasivoPrestamo");
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
