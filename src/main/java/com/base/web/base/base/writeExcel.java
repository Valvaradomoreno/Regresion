package com.base.web.base.base;

import java.io.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class writeExcel {
    private static XSSFCell Cell;

    public void writeToExcel(String dataToWrite,int rowCount, int ColNum) throws IOException{

        //Crea un objeto de la clase File para abrir el archivo xlsx

        String filePath = System.getProperty("user.dir")+"\\src\\Excel";
        String fileName = "AltaCuentaAhorro.xlsx";
        String sheetName = "AltaCuentaAhorro";
        File file =    new File(filePath+"\\"+fileName);

        //Cree un objeto de la clase FileInputStream para leer el archivo de Excel

        FileInputStream inputStream = new FileInputStream(file);

        Workbook workbook = null;

        //Encuentre la extensión del archivo dividiendo el nombre del archivo en subcadena y obteniendo solo el nombre de la extensión 
        String fileExtensionName = fileName.substring(fileName.indexOf("."));

        //Verifique la condición si el archivo es un archivo xlsx 
        if(fileExtensionName.equals(".xlsx")){
        //Si es un archivo xlsx, entonces cree un objeto de la clase 
        workbook = new XSSFWorkbook(inputStream);

        }
        else if(fileExtensionName.equals(".xls")){
            //Si es un archivo xlsx, entonces cree un objeto de la clase 
            workbook = new HSSFWorkbook(inputStream);

        }
         //Leer hoja de Excel por nombre   
         Sheet sheet = workbook.getSheet(sheetName);
       
        //Obtener el recuento actual de filas en el archivo de Excel 
        //int rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum();

        //Obtener la primera fila de la hoja 
        Row row = sheet.getRow(0);

        //Cree una nueva fila y agréguela al final de la hoja 
        Row newRow = sheet.createRow(rowCount+1);

        //Crear un bucle sobre la celda de la Fila recién creada
     //   for(int j = 0; j < row.getLastCellNum(); j++){

        //Rellenar datos en la fila
            //int ColNum = 7;
            //Cell cell = sheet.getRow(j).getCell(ColNum);
            Cell cell = sheet.getRow(rowCount+1).createCell(ColNum);
            cell.setCellValue(dataToWrite);
   // }

        //Close input stream
        inputStream.close();

        //Create an object of FileOutputStream class to create write data in excel file

        FileOutputStream outputStream = new FileOutputStream(file);

        //write data in the excel file

        workbook.write(outputStream);

        //close output stream

        outputStream.close();
        
    }

    public void writeToExcelTrans(String dataToWrite,int rowCount, int ColNum) throws IOException{

        //Crea un objeto de la clase File para abrir el archivo xlsx

        String filePath = System.getProperty("user.dir")+"\\src\\Excel";
        String fileName = "AltaCuentaAhorro.xlsx";
        String sheetName = "AltaCuentaAhorro";
        File file =    new File(filePath+"\\"+fileName);

        //Cree un objeto de la clase FileInputStream para leer el archivo de Excel

        FileInputStream inputStream = new FileInputStream(file);

        Workbook workbook = null;

        //Encuentre la extensión del archivo dividiendo el nombre del archivo en subcadena y obteniendo solo el nombre de la extensión
        String fileExtensionName = fileName.substring(fileName.indexOf("."));

        //Verifique la condición si el archivo es un archivo xlsx
        if(fileExtensionName.equals(".xlsx")){
            //Si es un archivo xlsx, entonces cree un objeto de la clase
            workbook = new XSSFWorkbook(inputStream);

        }
        else if(fileExtensionName.equals(".xls")){
            //Si es un archivo xlsx, entonces cree un objeto de la clase
            workbook = new HSSFWorkbook(inputStream);

        }
        //Leer hoja de Excel por nombre
        Sheet sheet = workbook.getSheet(sheetName);

        //Obtener el recuento actual de filas en el archivo de Excel
        //int rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum();

        //Obtener la primera fila de la hoja
        Row row = sheet.getRow(0);

        //Cree una nueva fila y agréguela al final de la hoja
        Row newRow = sheet.createRow(rowCount+1);

        //Crear un bucle sobre la celda de la Fila recién creada
        //for(int j = 0; j < row.getLastCellNum(); j++){

            //Rellenar datos en la fila
            //int ColNum = 7;
            //Cell cell = sheet.getRow(j).getCell(ColNum);
            Cell cell = sheet.getRow(rowCount+1).getCell(ColNum);
            cell.setCellValue(dataToWrite);
        //}

        //Close input stream

        inputStream.close();

        //Create an object of FileOutputStream class to create write data in excel file

        FileOutputStream outputStream = new FileOutputStream(file);

        //write data in the excel file

        workbook.write(outputStream);

        //close output stream

        outputStream.close();

    }

    public void writeToExcelFecha(String dataToWrite,int rowCount, int ColNum) throws IOException{

        //Crea un objeto de la clase File para abrir el archivo xlsx

        String filePath = System.getProperty("user.dir")+"\\src\\Excel";
        String fileName = "AltaCuentaAhorro.xlsx";
        String sheetName = "AltaCuentaAhorro";
        File file =    new File(filePath+"\\"+fileName);

        //Cree un objeto de la clase FileInputStream para leer el archivo de Excel

        FileInputStream inputStream = new FileInputStream(file);

        Workbook workbook = null;

        //Encuentre la extensión del archivo dividiendo el nombre del archivo en subcadena y obteniendo solo el nombre de la extensión
        String fileExtensionName = fileName.substring(fileName.indexOf("."));

        //Verifique la condición si el archivo es un archivo xlsx
        if(fileExtensionName.equals(".xlsx")){
            //Si es un archivo xlsx, entonces cree un objeto de la clase
            workbook = new XSSFWorkbook(inputStream);

        }
        else if(fileExtensionName.equals(".xls")){
            //Si es un archivo xlsx, entonces cree un objeto de la clase
            workbook = new HSSFWorkbook(inputStream);

        }
        //Leer hoja de Excel por nombre
        Sheet sheet = workbook.getSheet(sheetName);

        //Obtener el recuento actual de filas en el archivo de Excel
        //int rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum();

        //Obtener la primera fila de la hoja
        Row row = sheet.getRow(0);

        //Cree una nueva fila y agréguela al final de la hoja
        Row newRow = sheet.createRow(rowCount+1);

        //Crear un bucle sobre la celda de la Fila recién creada
        //for(int j = 0; j < row.getLastCellNum(); j++){

            //Rellenar datos en la fila
            //int ColNum = 7;
            //Cell cell = sheet.getRow(j).getCell(ColNum);
            Cell cell = sheet.getRow(rowCount+1).getCell(ColNum);
            cell.setCellValue(dataToWrite);
       // }

        //Close input stream

        inputStream.close();

        //Create an object of FileOutputStream class to create write data in excel file

        FileOutputStream outputStream = new FileOutputStream(file);

        //write data in the excel file

        workbook.write(outputStream);

        //close output stream

        outputStream.close();

    }
}