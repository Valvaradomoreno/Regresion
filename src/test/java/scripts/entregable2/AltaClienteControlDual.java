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
import java.util.concurrent.TimeUnit;

public class AltaClienteControlDual {

    WebDriver driver;
    public ExtentSparkReporter spark;
    public ExtentReports extent;
    public ExtentTest logger;

    AltaClienteSinBiometria altaClienteSinBiometria = new AltaClienteSinBiometria();
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
        ArrayList<String> profesion =readExcelData(12);
        ArrayList<String> ubigeo =readExcelData(13);
        ArrayList<String> celular =readExcelData(14);
        ArrayList<String> correo =readExcelData(15);
        ArrayList<String> usuario2= readExcelData(16);


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
                    driver.findElement(By.id("signOnName")).sendKeys(usuario.get(i));
                    driver.findElement(By.id("password")).sendKeys(contraseña.get(i));
                    driver.findElement(By.id("sign-in")).click();

                    // Then Redirecciona al Home Page
                    Thread.sleep(2000);
                    //Assert.assertEquals(exp_message, actual);
                    System.out.println("assert complete");

//                    WebElement iframe0 = driver.findElement(By.xpath("/html/frameset/frame[1]"));
//                    driver.switchTo().frame(iframe0);
//                    driver.findElement(By.id("commandValue")).sendKeys("CUSTOMER,INPUT.BRANCH.PE.BIOMETRY");
//                    driver.findElement(By.id("cmdline_img")).click();
//                    driver.switchTo().parentFrame();

                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/frameset/frame[2]")));
                    WebElement iframe1 = driver.findElement(By.xpath("/html/frameset/frame[2]"));
                    driver.switchTo().frame(iframe1);

                    driver.findElement(By.id("imgError")).click();

                    driver.findElement(By.xpath("//img[@alt='Cliente']")).click();

                    driver.findElement(By.xpath("//a[contains(text(),'Alta con Biometría ')]")).click();
                    driver.switchTo().parentFrame();

                    String MainWindow0=driver.getWindowHandle();
                    Set<String> s0=driver.getWindowHandles();
                    Iterator<String> i0=s0.iterator();

                    while(i0.hasNext())
                    {
                        String ChildWindow=i0.next();

                        if(!MainWindow0.equalsIgnoreCase(ChildWindow))
                        {
                            driver.switchTo().window(ChildWindow);
                        }
                    }

                    //driver.findElement(By.xpath("//img[@alt='New Deal']")).click();

                    wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:MNEMONIC")));
                    driver.findElement(By.id("fieldName:MNEMONIC")).sendKeys(mnemocino.get(i));
                    Thread.sleep(1000);
                    driver.findElement(By.id("fieldName:LEGAL.ID:1")).sendKeys(dni.get(i));
                    driver.findElement(By.id("fieldName:FAMILY.NAME")).click();
                    Thread.sleep(4000);
                    driver.findElement(By.id("fieldName:FAMILY.NAME")).sendKeys(apaterno.get(i));
                    driver.findElement(By.id("fieldName:NAME.2:1")).sendKeys(amaterno.get(i));
                    driver.findElement(By.id("fieldName:NAME.1:1")).sendKeys(nombre.get(i));
                    driver.findElement(By.id("fieldName:SHORT.NAME:1")).sendKeys(ncompleto.get(i));
                    Select selectProducto = new Select(driver.findElement(By.id("fieldName:MARITAL.STATUS")));
                    selectProducto.selectByVisibleText(estadocivil.get(i));
                    driver.findElement(By.id("radio:mainTab:GENDER")).click();
                    driver.findElement(By.id("fieldName:DATE.OF.BIRTH")).sendKeys(nacimiento.get(i));

                    driver.findElement(By.xpath("//img[@dropfield='fieldName:OCCUPATION:1']")).click();
                    Thread.sleep(2500);
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//td[contains(text(),'ABOGADO')]")));
                    driver.findElement(By.xpath("//td[contains(text(),'"+profesion.get(i)+"')]")).click();

                    driver.findElement(By.xpath("//img[@dropfield='fieldName:STREET:1']")).click();
                    Thread.sleep(3000);
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//td[contains(text(),'"+ubigeo.get(i)+"')]")));
                    driver.findElement(By.xpath("//td[contains(text(),'"+ubigeo.get(i)+"')]")).click();

                    driver.findElement(By.id("fieldName:ADDRESS:1:1")).sendKeys(gbdirec.get(i));

                    driver.findElement(By.xpath("//img[@dropfield='fieldName:PHONE.1:1']")).click();
                    Thread.sleep(2500);
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//td[contains(text(),'PERSONAL')]")));
                    driver.findElement(By.xpath("//td[contains(text(),'PERSONAL')]")).click();
                    driver.findElement(By.id("fieldName:SMS.1:1")).sendKeys(""+celular.get(i)+"");
                    driver.findElement(By.id("fieldName:EMAIL.1:1")).sendKeys(""+correo.get(i)+"");

                    String screenshotPath1 = getScreenShot(driver, "Fin del Caso");

                    driver.findElement(By.xpath("//span[contains(text(),'Datos Laborales')]")).click();
                    Thread.sleep(1000);

                    driver.findElement(By.id("fieldName:EMPLOYERS.NAME:1")).sendKeys(empresa.get(i));

                    driver.findElement(By.xpath("//span[contains(text(),'Proteccion de Datos')]")).click();
                    driver.findElement(By.id("radio:tab3:CONFID.TXT")).click();

                    driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();

                    Thread.sleep(5000);

                    driver.findElement(By.xpath("//img[@alt='Commit the deal']")).click();
                    Thread.sleep(3000);
                    driver.findElement(By.xpath("//a[contains(text(),'Autenticacion Biometrica')]")).click();
                    Thread.sleep(3000);
                    driver.findElement(By.xpath("//img[@alt='Commit the deal']")).click();
                    Thread.sleep(3000);


                    // CONTROL DUAL

                    driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
                    driver.get("https://10.167.21.100:8480/BrowserWebSAD/servlet/BrowserServlet?");

                    wait.until(ExpectedConditions.elementToBeClickable(By.id("signOnName")));

                    driver.findElement(By.id("signOnName")).sendKeys(usuario2.get(i));
                    driver.findElement(By.id("password")).sendKeys(contraseña.get(i));
                    driver.findElement(By.id("sign-in")).click();
                    Thread.sleep(1000);

                    WebElement iframe01 = driver.findElement(By.xpath("/html/frameset/frame[1]"));
                    driver.switchTo().frame(iframe01);
                    driver.findElement(By.id("commandValue")).sendKeys("ENQ BRIP.ENQ.DUAL.CONTROL");
                    driver.findElement(By.id("cmdline_img")).click();
                    driver.switchTo().parentFrame();


                    String MainWindow1=driver.getWindowHandle();
                    Set<String> s1=driver.getWindowHandles();
                    Iterator<String> i1=s1.iterator();

                    while(i1.hasNext())
                    {
                        String ChildWindow=i1.next();

                        if(!MainWindow1.equalsIgnoreCase(ChildWindow))
                        {
                            driver.switchTo().window(ChildWindow);
                        }
                    }


                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//label[contains(text(),'Número de Identificación')]")));
                    String attr1 = driver.findElement(By.xpath("//label[contains(text(),'Número de Identificación')]")).getAttribute("for");
                    driver.findElement(By.id(attr1)).clear();
                    driver.findElement(By.id(attr1)).sendKeys(dni.get(i));
                    driver.findElement(By.xpath("//a[@alt='Run Selection']")).click();
                    Thread.sleep(2000);

                    driver.findElement(By.xpath("//img[@alt='Autorizar']")).click();
                    Thread.sleep(2000);

                    driver.findElement(By.xpath("//img[@alt='Authorises a deal']")).click();
                    Thread.sleep(2000);

//                    if(driver.findElement(By.id("errorImg")).isDisplayed()){
//                        driver.findElement(By.id("errorImg")).click();
//                    }else{
//                        System.out.println("no hay");
//                    }

                    //Se muestra el codigo de alta
                    String cod = driver.findElement(By.xpath("//*[@id='messages']/tbody/tr[2]/td[2]/table[2]/tbody/tr/td")).getText();
                    String sSubCadena = cod.substring(22,39);
                    System.out.println(sSubCadena);
                    write(i+1, 18, sSubCadena);


                    String screenshotPath = getScreenShot(driver, "Fin del Caso");
                    logger.log(Status.PASS, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Fin del Caso", ExtentColor.GREEN));
                    extent.flush();
                    write(i+1, 17, "PASSED");

                    DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
                    String fecha = dateFormat.format(new Date());
                    System.out.println(fecha);
                    write(i+1, 19, fecha);

                    driver.quit();
                }

            }catch (Exception e){
                String screenshotPath = getScreenShot(driver, "Error");
                logger.log(Status.FAIL, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Error: "+e, ExtentColor.RED));
                extent.flush();
                write(i+1, 17, "FAILED");
                write(i+1, 18, "");

                DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
                String fecha = dateFormat.format(new Date());
                System.out.println(fecha);
                write(i+1, 19, fecha);
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
