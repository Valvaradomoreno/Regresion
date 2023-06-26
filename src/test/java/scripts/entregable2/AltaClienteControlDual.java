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

public class AltaClienteControlDual {

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
    public void AltaClienteControlDual()throws IOException, InterruptedException, AWTException {

        extent = new ExtentReports();
        spark = new ExtentSparkReporter(System.getProperty("user.dir") + "/test-output/reports2/AltaClienteControlDual/Report.html");
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
        ArrayList<String> mnemocino =readExcelData(2);
        ArrayList<String> dni =readExcelData(3);
        ArrayList<String> apaterno =readExcelData(4);
        ArrayList<String> amaterno =readExcelData(5);
        ArrayList<String> nombre =readExcelData(6);
        ArrayList<String> ncompleto =readExcelData(7);
        ArrayList<String> estadocivil =readExcelData(8);
        ArrayList<String> nacimiento =readExcelData(9);
        ArrayList<String> gbdirec =readExcelData(10);
        ArrayList<String> empresa =readExcelData(11);
        ArrayList<String> usuario2= readExcelData(12);


        int filas=usuario.size();
        for(int i=0;i<usuario.size();i++) {
            try {
                if(i<(filas)) {


                    System.out.println("-----------------------------------");
                    System.out.println("Nuevo Test " + i);
                    int caso = i+1;
                    logger = extent.createTest("Nuevo Test " + caso);

                    // ** DESDE AQUI EMPIEZA EL TEST




                    //  CONTROL DUAL **********************

                    driver = new ChromeDriver();
                    driver.manage().window().maximize();
                    driver.get("https://10.167.21.100:8480/BrowserWebSAD/servlet/BrowserServlet?");

                    Thread.sleep(1000);
                    driver.findElement(By.id("details-button")).click();
                    driver.findElement(By.id("proceed-link")).click();

                    // When El usuario ingresa el "<usuario>" y "<contraseña>"
                    Thread.sleep(5000);
                    WebDriverWait wait = new WebDriverWait(driver, 40);

                    wait.until(ExpectedConditions.elementToBeClickable(By.id("signOnName")));
                    driver.findElement(By.id("signOnName")).sendKeys(usuario2.get(i));
                    driver.findElement(By.id("password")).sendKeys(contraseña.get(i));
                    driver.findElement(By.id("sign-in")).click();

                    // Then Redirecciona al Home Page
                    Thread.sleep(2000);
                    //Assert.assertEquals(exp_message, actual);
                    System.out.println("assert complete");

                    WebElement iframe2 = driver.findElement(By.xpath("/html/frameset/frame[2]"));
                    driver.switchTo().frame(iframe2);
                    driver.findElement(By.id("imgError")).click();
                    driver.findElement(By.xpath("//img[@alt='Cliente']")).click();
                    driver.findElement(By.xpath("//a[contains(text(),'Control dual Biometria ')]")).click();

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
                    Thread.sleep(2000);

                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//label[contains(text(),'Número de Identificación')]")));
                    String attr1 = driver.findElement(By.xpath("//label[contains(text(),'Número de Identificación')]")).getAttribute("for");
                    driver.findElement(By.id(attr1)).clear();
                    driver.findElement(By.id(attr1)).sendKeys(dni.get(i));
                    driver.findElement(By.xpath("//a[@alt='Run Selection']")).click();
                    Thread.sleep(2000);

                    driver.findElement(By.xpath("//img[@alt='Autorizar']")).click();
                    Thread.sleep(2000);

                    driver.findElement(By.xpath("//img[@alt='Authorises a deal']")).click();

                    wait.until(ExpectedConditions.elementToBeClickable(By.id("errorImg")));
                    driver.findElement(By.id("errorImg")).click();

                    Thread.sleep(80000);


                    //Se muestra el codigo de alta
                    String cod = driver.findElement(By.xpath("//*[@id='messages']/tbody/tr[2]/td[2]/table[2]/tbody/tr/td")).getText();
                    String sSubCadena = cod.substring(22,39);
                    System.out.println(sSubCadena);
                    write(i+1, 14, sSubCadena);


                    String screenshotPath = getScreenShot(driver, "Fin del Caso");
                    logger.log(Status.PASS, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Fin del Caso", ExtentColor.GREEN));
                    extent.flush();
                    write(i+1, 13, "PASSED");

                    DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
                    String fecha = dateFormat.format(new Date());
                    System.out.println(fecha);
                    write(i+1, 15, fecha);

                    driver.quit();
                }

            }catch (Exception e){
                String screenshotPath = getScreenShot(driver, "Error");
                logger.log(Status.FAIL, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Error: "+e, ExtentColor.RED));
                extent.flush();
                write(i+1, 13, "FAILED");
                write(i+1, 14, "");

                DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
                String fecha = dateFormat.format(new Date());
                System.out.println(fecha);
                write(i+1, 15, fecha);
                System.out.println("Error: " + e);
                driver.quit();
            }
        }

    }

    public void write(int i, int celda, String dato) throws IOException {
        String path = System.getProperty("user.dir") + "/src/Excel/entregable2/AltaClienteControlDual.xlsx";
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
        String destination = System.getProperty("user.dir") + "/test-output/reports2/AltaClienteControlDual/Images/" + screenshotName + dateName + ".png";
        File finalDestination = new File(destination);
        FileUtils.copyFile(source, finalDestination);
        return destination;
    }

    public static ArrayList<String> readExcelData(int colNo) throws IOException {

        FileInputStream fis=new FileInputStream(System.getProperty("user.dir") + "/src/Excel/entregable2/AltaClienteControlDual.xlsx");
        XSSFWorkbook wb=new XSSFWorkbook(fis);
        XSSFSheet s=wb.getSheet("AltaClienteControlDual");
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
