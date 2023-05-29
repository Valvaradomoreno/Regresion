package scripts.entregable4;

import com.base.web.base.base.ThreadLocalDriver;
import io.cucumber.java.After;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import scripts.entregable2.*;

import java.awt.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class Entregable4 {

	CambioTasaDPF cambioTasaDPF = new CambioTasaDPF();
	CancelacionAnticipadaDPF cancelacionAnticipadaDPF = new CancelacionAnticipadaDPF();
	CondonacionDeuda condonacionDeuda = new CondonacionDeuda();
	CuotaComodin cuotaComodin = new CuotaComodin();
	FondeoDPF fondeoDPF = new FondeoDPF();
	RetiroCuentaCTS retiroCuentaCTS = new RetiroCuentaCTS();

	WebDriver driver;

	@BeforeTest
    public void startTest() {

    }
        
	@BeforeMethod
    public void openApplication() {
		System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir") + "/src/main/resources/drivers/chromedriver");
    }


	
	@Test
	public void entregable4()throws IOException, InterruptedException, AWTException {

		ArrayList<String> caso=readExcelData(0);
		ArrayList<String> ejecutar =readExcelData(1);

		int filas=caso.size();
  		for(int i=0;i<caso.size();i++) {
			System.out.println("caso get: "+caso.size());
  			if(i<(filas)){
				System.out.println("caso get2: "+caso.get(i));
				if((caso.get(i).equals("Cambio Tasa DPF")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("CambioTasaDPF");
					cambioTasaDPF.CambioTasaDPF();
				}if((caso.get(i).equals("Cancelacion Anticipada DPF")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("CancelacionAnticipadaDPF");
					cancelacionAnticipadaDPF.CancelacionAnticipadaDPF();
				}else if((caso.get(i).equals("Condonacion Deuda")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("CondonacionDeuda");
					condonacionDeuda.CondonacionDeuda();
				} else if((caso.get(i).equals("Cuota Comodin")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("CuotaComodin");
					cuotaComodin.CuotaComodin();
				} else if((caso.get(i).equals("Fondeo DPF")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("FondeoDPF");
					fondeoDPF.FondeoDPF();
				}else if((caso.get(i).equals("Retiro Cuenta CTS")) && (ejecutar.get(i).equals("Si"))) {
					System.out.println("RetiroCuentaCTS");
					retiroCuentaCTS.RetiroCuentaCTS();
				}
				else{
					System.out.println("NO SE EJECUTA");
					continue;
				}
			}

  		}

}	
		

	public static ArrayList<String> readExcelData(int colNo) throws IOException {
		FileInputStream fis=new FileInputStream(System.getProperty("user.dir") + "/src/Excel/entregable4/Inputs.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet s=wb.getSheet("Entregable4");
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
