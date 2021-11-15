package com.translation.tests;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteXLSXLoop {

    private static final File transFile= new File("D:\\xlsx\\Trans.xlsx");

    public static void main(String[] args) throws IOException {

        FileInputStream fisWriting = new FileInputStream(transFile);
        XSSFWorkbook workbook = new XSSFWorkbook (fisWriting);
        XSSFSheet sheet = workbook.getSheetAt(0);
        //Create rowNums Rows
        int rowNums=3;
        int colNums=5;
        XSSFRow[] row=new XSSFRow[rowNums];
        for (int i = 0; i <rowNums ; i++) {
             row[i]= sheet.createRow(i);

             //create colNums cells in the i row:
            XSSFCell[] cell=new XSSFCell[colNums];
            for (int j = 0; j < colNums; j++) {
                cell[j]=row[i].createCell(j);
                //assign a value to each cell[i,j]:
                cell[j].setCellValue("row"+i+"cell"+j);
            }

        }
        fisWriting.close();
        FileOutputStream fos =new FileOutputStream(transFile);
        workbook.write(fos);
        fos.close();
        System.out.println("Done: values are written in "+transFile);


    }

}
