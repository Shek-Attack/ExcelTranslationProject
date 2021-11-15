package com.translation.tests;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ReadXLSXLoop {

    public static void main(String[] args) throws IOException {
        FileInputStream fisReading = new FileInputStream("D:\\xlsx\\SourceSample.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook (fisReading);
        XSSFSheet sheet = workbook.getSheetAt(0);

        //reading from cells using for loop:
        int rowNums= sheet.getLastRowNum();
        int colNums=12;
        for (int i = 0; i <=rowNums; i++) {
            XSSFRow rowR=sheet.getRow(i);

            for (int j = 0; j < colNums; j++) {
                XSSFCell cellR=rowR.getCell(j);
                System.out.println("cellR("+i+","+j+") = " + cellR.toString());
            }
            System.out.println("============== End of Row "+i+" Read in ===================");// ending of a row
        }


        fisReading.close();

    }

}
