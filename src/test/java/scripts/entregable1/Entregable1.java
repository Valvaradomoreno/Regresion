package scripts.entregable1;

import com.base.web.base.base.ThreadLocalDriver;
import io.cucumber.java.After;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.awt.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class Entregable1 {

	AltaCuentaAhorro alta = new AltaCuentaAhorro();
	AltaCuentaCTS altaCTS = new AltaCuentaCTS();
	ActualizacionIntangibleCTS actualizacionIntangibleCTS= new ActualizacionIntangibleCTS();
	CambioSimpleAPlus cambioSimpleAPlus = new CambioSimpleAPlus();
	InactividadManualCuentaAhorros inactividadManualCuentaAhorros = new InactividadManualCuentaAhorros();


	BuscarCuenta buscar = new BuscarCuenta();
    WebDriver driver;

	@BeforeTest
    public void startTest() {

    }
        
	@BeforeMethod
    public void openApplication() {
		System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir") + "/src/main/resources/drivers/chromedriver");
    }


	
	@Test
	public void entregable1()throws IOException, InterruptedException, AWTException {

		ArrayList<String> caso=readExcelData(0);
		ArrayList<String> ejecutar =readExcelData(1);

		int filas=caso.size();
  		for(int i=0;i<caso.size();i++) {
			System.out.println("caso get: "+caso.size());
  			if(i<(filas)){
				System.out.println("caso get2: "+caso.get(i));
				if((caso.get(i).equals("Alta Cuenta Ahorro")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("Alta Cuenta Ahorro");
					alta.altaCuentaAhorro();
				}else if ((caso.get(i).equals("Buscar Cuenta")) && (ejecutar.get(i).equals("Si"))){
					System.out.println("Buscar Cuenta");
					buscar.buscarCuenta();
				}else if ((caso.get(i).equals("Alta Cuenta CTS")) && (ejecutar.get(i).equals("Si"))){
					System.out.println("Alta Cuenta CTS");
					altaCTS.altaCuentaCTS();
				}else if ((caso.get(i).equals("Actualizacion Intangible CTS")) && (ejecutar.get(i).equals("Si"))){
					System.out.println("Actualizacion Intangible CTS");
					actualizacionIntangibleCTS.ActualizacionIntangibleCTS();
				}else if ((caso.get(i).equals("Cambio Simple A Plus")) && (ejecutar.get(i).equals("Si"))){
					System.out.println("Cambio Simple A Plus");
					cambioSimpleAPlus.CambioSimpleAPlus();
				}else if ((caso.get(i).equals("Inactividad Manual Cuenta Ahorros")) && (ejecutar.get(i).equals("Si"))){
					System.out.println("Inactividad Manual Cuenta Ahorros");
					inactividadManualCuentaAhorros.InactividadManualCuentaAhorros();
				}else{
					System.out.println("NO SE EJECUTA");
					continue;
				}
			}

  		}

}	
		

	public static ArrayList<String> readExcelData(int colNo) throws IOException {
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir") + "/src/Excel/Inputs.xlsx");

		//FileInputStream fis=new FileInputStream("D:\\Proyectos\\Regresion\\src\\Excel\\Inputs.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet s=wb.getSheet("Entregable1");
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
