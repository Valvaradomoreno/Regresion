package scripts.entregable3;

import com.base.web.base.base.ThreadLocalDriver;
import io.cucumber.java.After;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import scripts.entregable1.CambioTasaCuentaAhorro;

import java.awt.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class Entregable3 {

	ActualizacionIntangibleCTS actualizacionIntangibleCTS = new ActualizacionIntangibleCTS();
	AjusteBILL ajusteBill = new AjusteBILL();
	AltaPrestamo altaPrestamo = new AltaPrestamo();
	CambioFechaPago cambioFechaPago = new CambioFechaPago();
	CambioT24aCastigo cambioT24aCastigo = new CambioT24aCastigo();
	CambioTasaCuentaAhorro cambioTasaCuentaAhorro = new CambioTasaCuentaAhorro();
	CancelacionPrestamo cancelacionPrestamo = new CancelacionPrestamo();
	CancelacionPrestamoAhorros cancelacionPrestamoAhorros = new CancelacionPrestamoAhorros();
	CuentaPagarSoles cuentaPagarSoles = new CuentaPagarSoles();
	CuentaPagarDolares cuentaPagarDolares = new CuentaPagarDolares();
	PagoAnticipadoPlazo pagoAnticipadoPlazo = new PagoAnticipadoPlazo();
	PagoAnticipadoValor pagoAnticipadoValor = new PagoAnticipadoValor();
	PagoCuotaPrestamo pagoCuotaPrestamo = new PagoCuotaPrestamo();
	PagoCuotaPrestamoCargoCuenta pagoCuotaPrestamoCargoCuenta = new PagoCuotaPrestamoCargoCuenta();

	WebDriver driver;

	@BeforeTest
    public void startTest() {

    }
        
	@BeforeMethod
    public void openApplication() {
		System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir") + "/src/main/resources/drivers/chromedriver");
    }


	
	@Test
	public void entregable3()throws IOException, InterruptedException, AWTException {

		ArrayList<String> caso=readExcelData(0);
		ArrayList<String> ejecutar =readExcelData(1);

		int filas=caso.size();
  		for(int i=0;i<caso.size();i++) {
			System.out.println("caso get: "+caso.size());
  			if(i<(filas)){
				System.out.println("caso get2: "+caso.get(i));
				if((caso.get(i).equals("Actualizacion Intangible CTS")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("ActualizacionIntangibleCTS");
					actualizacionIntangibleCTS.ActualizacionIntangibleCTS();
				}if((caso.get(i).equals("Ajuste Bill")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("Ajuste Bill");
					ajusteBill.AjusteBill();
				}else if((caso.get(i).equals("Alta Prestamo")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("AltaPrestamo");
					altaPrestamo.AltaPrestamo();
				} else if((caso.get(i).equals("Cambio Fecha Pago")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("CambioFechaPago");
					cambioFechaPago.CambioFechaPago();
				} else if((caso.get(i).equals("CambioT24aCastigo")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("CambioT24aCastigo");
					cambioT24aCastigo.CambioT24aCastigo();
				}else if((caso.get(i).equals("Cambio Tasa Cuenta Ahorro")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("CambioTasaCuentaAhorro");
					cambioTasaCuentaAhorro.CambioTasaCuentaAhorro();
				}else if((caso.get(i).equals("Cancelacion Prestamo")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("CancelacionPrestamo");
					cancelacionPrestamo.CancelacionPrestamo();
				}else if((caso.get(i).equals("Cancelacion Prestamo Ahorros")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("CancelacionPrestamoAhorros");
					cancelacionPrestamoAhorros.CancelacionPrestamoAhorros();
				}else if((caso.get(i).equals("Cuenta Pagar Dolares")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("CuentaPagarDolares");
					cuentaPagarDolares.CuentaPagarDolares();
				}else if((caso.get(i).equals("Cuenta Pagar Soles")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("CuentaPagarSoles");
					cuentaPagarSoles.CuentaPagarSoles();
				}else if((caso.get(i).equals("Pago Anticipado Plazo")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("PagoAnticipadoPlazo");
					pagoAnticipadoPlazo.PagoAnticipadoPlazo();
				}else if((caso.get(i).equals("Pago Anticipado Valor")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("PagoAnticipadoValor");
					pagoAnticipadoValor.PagoAnticipadoValor();
				}else if((caso.get(i).equals("Pago Cuota Prestamo")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("PagoCuotaPrestamo");
					pagoCuotaPrestamo.PagoCuotaPrestamo();
				}else if((caso.get(i).equals("Pago Cuota Prestamo Cargo Cuenta")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("PagoCuotaPrestamoCargoCuenta");
					pagoCuotaPrestamoCargoCuenta.PagoCuotaPrestamoCargoCuenta();
				}
				else{
					System.out.println("NO SE EJECUTA");
					continue;
				}
			}

  		}

}	
		

	public static ArrayList<String> readExcelData(int colNo) throws IOException {
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir") + "/src/Excel/entregable3/Inputs.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet s=wb.getSheet("Entregable3");
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
