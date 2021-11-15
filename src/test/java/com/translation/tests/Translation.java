package com.translation.tests;

import com.translation.pages.TranslationTestPage;
import com.translation.utilities.BrowserUtils;
import com.translation.utilities.Driver;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.openqa.selenium.Keys;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;

public class Translation {

   TranslationTestPage translationTestPage=new TranslationTestPage();

    // translate from
    String source="chinese"; // can be any language supported by Google Translate
    // to
    String[] target={"english", "german", "french", "russian", "turkish", "japanise", "arabic", "spanish"};

    @Test
    public void translationTest() throws IOException {

        // read from the Excel file:
        //Path of the Excel file
        String pathSource = "SourceSampleTable.xlsx";
        FileInputStream fs = new FileInputStream(pathSource);
        //Creating a workbook
        XSSFWorkbook workbook = new XSSFWorkbook(fs);
        XSSFSheet sheetSource = workbook.getSheetAt(0);

        int rowNum = sheetSource.getLastRowNum();
        int colNum = 12; // starting 0,1,2,...,11: so it is 12 columns

        System.out.println("rowNum = " + rowNum);
        System.out.println("colNum = " + colNum);

/*
 * ************************************************************************* */
        // write to the Excel file:
        //Path of the Excel file
        String pathTarget = "TargetSample.xlsx";
        FileInputStream fsTarget = new FileInputStream(pathTarget);
        //Creating a workbook
        XSSFWorkbook workbookTarget = new XSSFWorkbook(fsTarget);
        XSSFSheet sheetTarget = workbookTarget.getSheetAt(0);



        // rows:

        String targetWord;
        for (int j = 1; j <2; j++) {
            System.out.print("");

            // go to google translate from "source" to "target":
            Driver.getDriver().get("https://www.deepl.com/translator#zh/en/");

            // columns of a row:
            for (int i = 0; i <colNum; i++) {

                // write into the sourceTextArea

                translationTestPage.sourceTextArea.sendKeys(Keys.CLEAR);
                translationTestPage.sourceTextArea.sendKeys(sheetSource.getRow(j).getCell(i)+"|", Keys.ENTER);
                System.out.println("sheetSource.getRow("+j+").getCell("+i+") = " + sheetSource.getRow(j).getCell(i));
                // source reading is complete

            }

            //read the translation from targetTextArea
            // wait for translation
            //wait just long enough depending on the length of the last cell.
            //wait just long enough depending on the length of the last cell.
            //int waitTimeForLongChar2Translate=1+(cellRContent.length()/1000); //wait 1 sec for 1k chars 2 translate
            //System.out.println("waitTimeForLongChar2Translate = " + waitTimeForLongChar2Translate);
            //BrowserUtils.wait(waitTimeForLongChar2Translate);

            System.out.println("");


            targetWord=translationTestPage.transTextArea.getText();
            System.out.println(targetWord);

           // translation is done as one row string, extract the cell values:
                // problem is below code:

                String[] entryRowCells=targetWord.split("\\|");
                String entryRowArray=Arrays.toString(entryRowCells);
            System.out.println("entryRowArray = " + entryRowArray);


            for (int i = 0; i <12; i++) {
                    XSSFCell cell = sheetTarget.getRow(j).createCell(i);

                    cell.setCellValue(Arrays.toString(entryRowCells));
                    System.out.println("Cell["+i+"] = " + entryRowCells[i]);

                }


            }

        System.out.println(" ");


        FileOutputStream fos = new FileOutputStream(pathTarget);
        workbookTarget.write(fos);
        fos.close();


        Driver.closeDriver();


    }




}
