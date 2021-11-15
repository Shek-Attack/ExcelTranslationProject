package com.translation.tests;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.*;

public class ReadTransLateWriteXLSX {

    // file to d=read from:
    private static final File sourceFile= new File("D:\\xlsx\\SourceSample.xlsx");
    //String sourceFile = "SourceSample.xlsx"; // if the file is directly under the project

    // file to write into:
    private static final File transFile= new File("D:\\xlsx\\Trans.xlsx");

    //obtaining row numbers from source sheet:
    //int rowNums= sheetR.getLastRowNum(); // for real job
    int rowNums=3; //0,1, ... // for testing only
    int colNums=12; // 0,1,2, ...,11 // for test and real job

    public String[][] cell2W2DArray=new String[rowNums+1][colNums];

    // translate from
    String source="chinese"; // can be any language supported by Google Translate
    // to
    String[] target={"english", "german", "french", "russian", "turkish", "japanese", "arabic", "spanish"};

    @Test
    public void translationTest() throws IOException {

        // For reading, need to define source for FileInputStream, workbook, sheet
        FileInputStream fisReading = new FileInputStream(sourceFile);
        XSSFWorkbook workbookR = new XSSFWorkbook(fisReading);
        XSSFSheet sheetR = workbookR.getSheetAt(0);

        // For writing, need to define source for FileInputStream, workbook, sheet
        FileInputStream fisWriting = new FileInputStream(transFile);
        XSSFWorkbook workbookW = new XSSFWorkbook (fisWriting);
        XSSFSheet sheetW = workbookW.getSheetAt(0);

        XSSFRow[] row=new XSSFRow[rowNums];
        for (int i = 0; i <rowNums ; i++) {
            // get rows from the sheetR as reading in info:
            XSSFRow rowR=sheetR.getRow(i);
            Row rowW=sheetW.getRow(i);

            // create rows from the sheetW to write
            row[i]= sheetW.createRow(i);

            //create colNums cells in the i row:
            XSSFCell[] cell=new XSSFCell[colNums];
            for (int j = 0; j < colNums; j++) {
                cell[j]=row[i].createCell(j);




                //assign a value to each cell[i,j]:
                cell[j].setCellValue("row"+i+"cell"+j); // assigning value
                cell2W2DArray[i][j]=cell[j].toString();
                System.out.println("cell2W2DArray["+i+"]["+j+"] = " + cell2W2DArray[i][j]);
            }// columns of a row ends here

        } //rows of an Excel table ends here


        // for writing the file:
        fisWriting.close();
        FileOutputStream fos =new FileOutputStream(transFile);
        workbookW.write(fos);
        fos.close();
        System.out.println("Done: values are written in "+transFile);


    }

}
