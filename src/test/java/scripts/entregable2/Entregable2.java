package scripts.entregable2;

import com.base.web.base.base.ThreadLocalDriver;
import io.cucumber.java.After;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.awt.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class Entregable2 {

	AltaClienteSinBiometria altaCliente= new AltaClienteSinBiometria();
	AltaClienteControlDual altaClienteControlDual = new AltaClienteControlDual();
	ModificarCliente modificarCliente= new ModificarCliente();
	ModificarClienteCorporativo modificarClienteCorporativo = new ModificarClienteCorporativo();
	PagoMasivoPrestamo pagoMasivoPrestamo = new PagoMasivoPrestamo();
	PagoTCyReversa pagoTCyReversa = new PagoTCyReversa();
	ProcesoAbonoMasivoCTS procesoAbonoMasivoCTS = new ProcesoAbonoMasivoCTS();
	ProcesoActualizacionDatosMasivos procesoActualizacionDatosMasivos = new ProcesoActualizacionDatosMasivos();
    WebDriver driver;

	@BeforeTest
    public void startTest() {

    }
        
	@BeforeMethod
    public void openApplication() {
		System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir") + "/src/main/resources/drivers/chromedriver");
    }


	
	@Test
	public void entregable2()throws IOException, InterruptedException, AWTException {

		ArrayList<String> caso=readExcelData(0);
		ArrayList<String> ejecutar =readExcelData(1);

		int filas=caso.size();
  		for(int i=0;i<caso.size();i++) {
			System.out.println("caso get: "+caso.size());
  			if(i<(filas)){
				System.out.println("caso get2: "+caso.get(i));
				if((caso.get(i).equals("Alta Cliente")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("Alta Cliente");
					altaCliente.AltaClienteSinBiometria();
				}if((caso.get(i).equals("Alta Cliente Control Dual")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("Alta Cliente Control Dual");
					altaClienteControlDual.AltaClienteControlDual();
				}else if((caso.get(i).equals("Modificar Cliente")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("ModificarCliente");
					modificarCliente.ModificarCliente();
				} else if((caso.get(i).equals("Modificar Cliente Corporativo")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("ModificarClienteCorporativo");
					modificarClienteCorporativo.ModificarClienteCorporativo();
				} else if((caso.get(i).equals("Pago Masivo Prestamo")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("PagoMasivoPrestamo");
					pagoMasivoPrestamo.PagoMasivoPrestamo();
				}else if((caso.get(i).equals("Pago TC Reversa")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("PagoTCyReversa");
					pagoTCyReversa.PagoTCyReversa();
				}else if((caso.get(i).equals("Proceso Abono Masivo CTS")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("ProcesoAbonoMasivoCTS");
					procesoAbonoMasivoCTS.ProcesoAbonoMasivoCTS();
				}else if((caso.get(i).equals("Proceso Actualizacion Datos Masivos")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("ProcesoActualizacionDatosMasivos");
					procesoActualizacionDatosMasivos.ProcesoActualizacionDatosMasivos();
				}
				else{
					System.out.println("NO SE EJECUTA");
					continue;
				}
			}

  		}

}	
		

	public static ArrayList<String> readExcelData(int colNo) throws IOException {
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir") + "/src/Excel/entregable2/Inputs.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet s=wb.getSheet("Entregable2");
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
