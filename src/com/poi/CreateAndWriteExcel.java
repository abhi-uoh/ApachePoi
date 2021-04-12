package com.poi;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class CreateAndWriteExcel {

    public static void main(String[] args) throws Exception {

        HSSFWorkbook workbook1 = new HSSFWorkbook();
        HSSFSheet sheet1 = workbook1.createSheet("DataFile"); //sheet name

        //row 0
        sheet1.createRow(0);
        sheet1.getRow(0).createCell(0).setCellValue("Hello");
        sheet1.getRow(0).createCell(1).setCellValue("World");

        //row1
        sheet1.createRow(1);
        sheet1.getRow(1).createCell(0).setCellValue("Jocata");
        sheet1.getRow(1).createCell(1).setCellValue("Roxana");


        File file = new File("/Users/abhi/Desktop/Programs/demo1/ApachePoi/datafiles/data3.xls");
        FileInputStream fos = new FileInputStream(file);
        workbook1.write(file);
        workbook1.close();
    }
}
