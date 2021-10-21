package com.cybertek.translation;

import com.cybertek.pages.TranslationTestPage;
import com.cybertek.utilities.BrowserUtils;
import com.cybertek.utilities.Driver;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.openqa.selenium.Keys;

import java.io.*;
import java.util.Arrays;

public class ReadTranslateWriteXLSX {

    TranslationTestPage translationTestPage=new TranslationTestPage();

    private static final File sourceFile= new File("D:\\xlsx\\SourceSample.xlsx");
    //String sourceFile = "SourceSample.xlsx"; // if the file is directly under the project
    private static final File transFile= new File("D:\\xlsx\\Trans.xlsx");

    @Test
    public void setTranslationTest() throws IOException {

        // For reading, need to define source for FileInputStream, workbook, sheet
        FileInputStream fisReading = new FileInputStream(sourceFile);
        XSSFWorkbook workbookR = new XSSFWorkbook (fisReading);
        XSSFSheet sheetR = workbookR.getSheetAt(0);

        // For writing, need to define source for FileInputStream, workbook, sheet
        FileInputStream fisWriting = new FileInputStream(transFile);
        XSSFWorkbook workbookW = new XSSFWorkbook (fisWriting);
        XSSFSheet sheetW = workbookW.getSheetAt(0);

        //reading from cells using for loop:
        //int rowNums= sheetR.getLastRowNum();
        int rowNums=2;
        int colNums=12;

        String translation;

        for (int i = 0; i <=rowNums; i++) {
            // get rows from the sheet
            XSSFRow rowR=sheetR.getRow(i);
            // create a row to write to the sheet
            XSSFRow rowW= sheetW.createRow(i);

            // go to DeppL to translate from "source" to "target":
            Driver.getDriver().get("https://www.deepl.com/translator#zh/en/");

            for (int j = 0; j < colNums; j++) {


                // get cell from the row
                XSSFCell cellR=rowR.getCell(j);// read/get content from cell(j)
                System.out.println("cellR("+i+","+j+") = " + cellR.toString()); // see what is there

                //translation begins here==============================

                //read cell from the source file into the sourceTextArea of DeepL
                translationTestPage.sourceTextArea.sendKeys(sheetR.getRow(i).getCell(j)+",  ", Keys.ENTER);
                // source reading is complete


                //create colNums cells in the i row:
                XSSFCell cellW=rowW.createCell(j); // cell for writing the content
                //assign a value cellR to cellW
                cellW.setCellValue(cellR.toString());


            }// End of a row

            BrowserUtils.wait(3);
            BrowserUtils.waitForVisibility(translationTestPage.transTextArea,5);
            //read the translation from targetTextArea and wait for translation to complete
            BrowserUtils.waitForClickability(translationTestPage.transTextArea, 15);
            translation=translationTestPage.transTextArea.getText();
            System.out.println("translation = " + translation);
            String[] translatedCells=translation.split(",  ");
            System.out.println("translatedCells = " + Arrays.toString(translatedCells));
            //translation ends here =================================

            System.out.println("============== End of Row "+i+" Read in ===================");

            fisWriting.close();
            FileOutputStream fos =new FileOutputStream(transFile);
            workbookW.write(fos);
            fos.close();
            System.out.println("Done: values are written in "+transFile);


        }// End of Excel Table

        fisReading.close();


    } // End of Test


}
