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

public class AltaPrestamo {

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
    public void AltaPrestamo()throws IOException, InterruptedException, AWTException {

        extent = new ExtentReports();
        spark = new ExtentSparkReporter(System.getProperty("user.dir") + "/test-output/reports3/AltaPrestamo/Report.html");
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
        ArrayList<String> documento =readExcelData(2);
        ArrayList<String> tipo_producto =readExcelData(3);
        ArrayList<String> ejecutivo =readExcelData(4);
        ArrayList<String> monto =readExcelData(5);
        ArrayList<String> fechamaduracion =readExcelData(6);
        ArrayList<String> tarifa1 =readExcelData(7);
        ArrayList<String> usuario2= readExcelData(8);
        ArrayList<String> cuenta= readExcelData(9);


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



                    //El usuario da click en Catalogo de Productos
                    Thread.sleep(1000);
                    driver.findElement(By.xpath("//a[contains(text(),'Catálogo de Productos ')]")).click();
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
                    driver.manage().window().maximize();
                    WebElement iframe3 = driver.findElement(By.xpath("/html/frameset/frameset[2]/frameset[1]/frame[2]"));
                    driver.switchTo().frame(iframe3);
                    driver.findElement(By.id("treestop9")).click();
                    driver.findElement(By.xpath("//*[@id='r9']/td[4]/a/img")).click();
                    driver.switchTo().parentFrame();

                    String screenshotPath1 = getScreenShot(driver, "Fin del Caso");

                    //El usuario crea un prestamo ya
                    WebElement iframe4 = driver.findElement(By.xpath("/html/frameset/frameset[2]/frameset[2]/frame[2]"));
                    driver.switchTo().frame(iframe4);
                    if(tipo_producto.get(i).equals("LD011005")){
                        driver.findElement(By.xpath("//*[@id='r2']/td[3]/a/img")).click();
                    }else if(tipo_producto.get(i).equals("LD011006")){
                        driver.findElement(By.xpath("//*[@id='r3']/td[3]/a/img")).click();
                    }else if(tipo_producto.get(i).equals("LD011011")){
                        driver.findElement(By.xpath("//*[@id='r7']/td[3]/a/img")).click();
                    }else if(tipo_producto.get(i).equals("LD011010")){
                        driver.findElement(By.xpath("//*[@id='r6']/td[3]/a/img")).click();
                    }else if(tipo_producto.get(i).equals("LD011002")){
                        driver.findElement(By.xpath("//*[@id='r1']/td[3]/a/img")).click();
                    }
                    driver.switchTo().parentFrame();

                    //El usuario ingresa documento "<documento>" y moneda
                    String MainWindow2=driver.getWindowHandle();
                    Set<String> s2=driver.getWindowHandles();
                    Iterator<String> i2=s2.iterator();

                    while(i2.hasNext())
                    {
                        String ChildWindow=i2.next();

                        if(!MainWindow.equalsIgnoreCase(ChildWindow))
                        {
                            driver.switchTo().window(ChildWindow);
                        }
                    }

                    driver.manage().window().maximize();


                    driver.findElement(By.id("fieldName:CUSTOMER:1")).sendKeys(documento.get(i));
                    driver.findElement(By.id("fieldName:CURRENCY")).sendKeys("PEN");

                    //El usuario prevalida
                   // wait.until(ExpectedConditions.elementToBeClickable(By.xpath("fieldName:AMOUNT.LOCAL.1:1"))).sendKeys(monto.get(i));
                    driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();

                    //El usuario ingresa datos de cliente "<ejecutivo>", marca virtual y personal para prestamo
                    wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:PRIMARY.OFFICER")));
                    driver.findElement(By.id("fieldName:PRIMARY.OFFICER")).sendKeys(ejecutivo.get(i));
                    driver.findElement(By.xpath("/html/body/div[5]/fieldset[3]/div/div/form[1]/div[3]/table/tbody/tr[2]/td/table/tbody/tr[8]/td[3]/table/tbody/tr/td[2]/input")).click();
                    driver.findElement(By.xpath("/html/body/div[5]/fieldset[3]/div/div/form[1]/div[3]/table/tbody/tr[2]/td/table/tbody/tr[9]/td[3]/table/tbody/tr/td[2]/input")).click();

                    String screenshotPath2 = getScreenShot(driver, "");

                    //El usuario captura la cuenta prestamo
                    String cod = driver.findElement(By.id("disabled_ACCOUNT.REFERENCE")).getText();
                    System.out.println("CUENTA : " +cod);
                    write(i+1, 10, cod);
                    write(i+1, 8, cod);

                    String arreglo = driver.findElement(By.id("disabled_ARRANGEMENT")).getText();
                    System.out.println("ARREGLO : " +arreglo);

                    //El usuario ingresa monto "<monto>"
                    driver.findElement(By.id("fieldName:AMOUNT")).sendKeys(monto.get(i));

                    //El usuario ingresa fechamaduracion "<fechamaduracion>"
                    driver.findElement(By.id("fieldName:MATURITY.DATE")).sendKeys(fechamaduracion.get(i));

                    //El usuario ingresa tasas "<tarifa1>"
                    driver.findElement(By.id("fieldName:FIXED.RATE:1")).sendKeys(tarifa1.get(i));
                    driver.findElement(By.xpath("/html/body/div[5]/fieldset[7]/div/div[3]/div/div/form[1]/div[3]/table/tbody/tr[2]/td/table/tbody/tr[1]/td[3]/input")).sendKeys(tarifa1.get(i));

                    //El usuario ingresa fechastart "<fechamaduracion>"
                    driver.findElement(By.id("fieldName:START.DATE:2:1")).sendKeys(fechamaduracion.get(i));

                    //El usuario prevalida
                    driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();

                    String screenshotPath5 = getScreenShot(driver, "Fin del Caso");

                    //El usuario hace commit
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']")));
                    driver.findElement(By.xpath("//img[@alt='Commit the deal']")).click();

                    //El usuario marca recibido y acepta prestamo "<documento>"
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div[3]/div/div[2]/form[1]/div[3]/table/tbody/tr[1]/td/table/tbody/tr/td[3]/select")));
                    Select selectProducto = new Select(driver.findElement(By.xpath("/html/body/div[3]/div/div[2]/form[1]/div[3]/table/tbody/tr[1]/td/table/tbody/tr/td[3]/select")));
                    selectProducto.selectByVisibleText("RECEIVED");

                    //El usuario hace commit
                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']")));
                    driver.findElement(By.xpath("//img[@alt='Commit the deal']")).click();

                    //Se muestra el codigo de transacción de alta prestamo
                    String cod2 = driver.findElement(By.xpath("//*[@id='messages']/tbody/tr[2]/td[2]/table[2]/tbody/tr/td")).getText();
                    String sSubCadena = cod2.substring(22,39);
                    System.out.println(sSubCadena);
                    write(i+1, 11, sSubCadena);


                    ////// DESEMBOLSO ******************


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
                    driver.findElement(By.xpath("//img[@alt='Transacciones de Préstamo']")).click();
                    driver.findElement(By.xpath("//*[@id=\"pane_\"]/ul/li/ul/li[4]/ul/li[7]/ul/li/span/text()")).click();

                    driver.findElement(By.xpath("//a[contains(text(),'AA Desembolso ')]")).click();
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

                    wait.until(ExpectedConditions.elementToBeClickable(By.id("fieldName:ACCOUNT.1:1")));
                    driver.findElement(By.id("fieldName:ACCOUNT.1:1")).sendKeys(cod);
                    driver.findElement(By.id("fieldName:AMOUNT.LOCAL.1:1")).sendKeys(monto.get(i));

                    String screenshotPath3 = getScreenShot(driver, "Fin del Caso");

                    driver.findElement(By.xpath("//img[@alt='Validate a deal']")).click();


                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt='Commit the deal']")));
                    driver.findElement(By.xpath("//img[@alt='Commit the deal']")).click();

                    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='messages']/tbody/tr[2]/td[2]/table[2]/tbody/tr/td"))).click();


                    String screenshotPath4 = getScreenShot(driver, "Fin del Caso");

                    logger.log(Status.PASS, MarkupHelper.createLabel("Tipos de Prestamo", ExtentColor.GREEN));
                    logger.log(Status.PASS,"Tipos de Prestamo", MediaEntityBuilder.createScreenCaptureFromPath(screenshotPath1).build());
                    logger.log(Status.PASS, MarkupHelper.createLabel("Estado de la cuenta", ExtentColor.GREEN));
                    logger.log(Status.PASS,"Estado de la cuenta", MediaEntityBuilder.createScreenCaptureFromPath(screenshotPath2).build());
                    logger.log(Status.PASS, MarkupHelper.createLabel("Desembolso", ExtentColor.GREEN));
                    logger.log(Status.PASS,"Desembolso", MediaEntityBuilder.createScreenCaptureFromPath(screenshotPath3).build());
                    logger.log(Status.PASS, MarkupHelper.createLabel("Alta creada", ExtentColor.GREEN));
                    logger.log(Status.PASS,"Alta creada", MediaEntityBuilder.createScreenCaptureFromPath(screenshotPath4).build());

                    extent.flush();
                    write(i+1, 10, "PASSED");

                    DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
                    String fecha = dateFormat.format(new Date());
                    System.out.println(fecha);
                    write(i+1, 12, fecha);

                    driver.quit();
                }

            }catch (Exception e){
                String screenshotPath = getScreenShot(driver, "Error");
                logger.log(Status.FAIL, MarkupHelper.createLabel(logger.addScreenCaptureFromPath(screenshotPath) + " Error: "+e, ExtentColor.RED));
                extent.flush();
                write(i+1, 10, "FAILED");
                write(i+1, 11, "");

                DateFormat dateFormat = new SimpleDateFormat("d MMM yyyy, HH:mm:ss");
                String fecha = dateFormat.format(new Date());
                System.out.println(fecha);
                write(i+1, 12, fecha);
                System.out.println("Error: " + e);
                driver.quit();
            }
        }

    }

    public void write(int i, int celda, String dato) throws IOException {
        String path = System.getProperty("user.dir") + "/src/Excel/entregable3/AltaPrestamo.xlsx";
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
        String destination = System.getProperty("user.dir") + "/test-output/reports3/AltaPrestamo/Images/" + screenshotName + dateName + ".png";
        File finalDestination = new File(destination);
        FileUtils.copyFile(source, finalDestination);
        return destination;
    }

    public static ArrayList<String> readExcelData(int colNo) throws IOException {

        FileInputStream fis=new FileInputStream(System.getProperty("user.dir") + "/src/Excel/entregable3/AltaPrestamo.xlsx");
        XSSFWorkbook wb=new XSSFWorkbook(fis);
        XSSFSheet s=wb.getSheet("AltaPrestamo");
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
