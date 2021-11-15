package com.translation.tests;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadXLSX {
    public static void main(String[] args) throws IOException {
        FileInputStream fisReading = new FileInputStream("D:\\xlsx\\SourceSample.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook (fisReading);
        XSSFSheet sheet = workbook.getSheetAt(0);


        //reading from cells using Iterator:
        Iterator<Row> ite = sheet.rowIterator(); // ite: rows
        while(ite.hasNext()){
            Row row = ite.next();
            Iterator<Cell> cite = row.cellIterator();
            while(cite.hasNext()){
                XSSFCell c = (XSSFCell) cite.next();
                System.out.print(c.toString() +" || ");
            }
            System.out.println();// ending of a row
        } //ending of all rows


        fisReading.close();

    }
}
