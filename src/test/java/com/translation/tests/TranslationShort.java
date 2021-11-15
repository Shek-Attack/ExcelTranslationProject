package com.translation.tests;

import com.translation.pages.TranslationTestPage;
import com.translation.utilities.BrowserUtils;
import com.translation.utilities.Driver;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.openqa.selenium.Keys;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

// more than couple of lines of excel table or more than 3900 chars,
// google translate refuses to do the translation
// so this attempt has only demo value, nothing more, unfortunately.

public class TranslationShort {

   TranslationTestPage translationTestPage=new TranslationTestPage();

    // translate from
    String source="chinese"; // can be any language supported by Google Translate
    // to
    String[] target={"english", "german", "french", "russian", "turkish", "japanise", "arabic", "spanish"};

    @Test
    public void setTranslationTest() throws IOException {

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
        XSSFWorkbook workbookW = new XSSFWorkbook(fsTarget);
        XSSFSheet sheetW = workbookW.getSheetAt(0);



        // rows:

        for (int j = 0; j <2; j++) {
            System.out.print("");

            Row rowR = sheetSource.getRow(j);

            // create a row to write to the sheet
            XSSFRow rowW= sheetW.createRow(j);

            // columns of a row:
            for (int i = 0; i < colNum; i++) {

                // go to DeepL translate from "source" to "target":
                //Driver.getDriver().get("https://www.google.com/search?q=" + source + "+to+" + target[0] + "+translate&oq=chinese+to+english&aqs=chrome.1.69i57j35i39j0i512l8.9361j0j15&sourceid=chrome&ie=UTF-8");
                Driver.getDriver().get("https://www.deepl.com/translator#zh/en/");

                // write into the sourceTextArea
                String sourceCell=sheetSource.getRow(j).getCell(i).toString();
                BrowserUtils.wait(1+sourceCell.length()/1000);
                translationTestPage.sourceTextArea.sendKeys(sourceCell + "  ", Keys.ENTER);
                //System.out.println("sheetSource.getRow(" + j + ").getCell(" + i + ") = " + sheetSource.getRow(j).getCell(i));
                System.out.println("sheetSource.getRow(" + j + ").getCell(" + i + ") = " + sourceCell);
                // source writing and reading is complete


                //get the translation from translationTextArea
                // wait for translation
                BrowserUtils.wait(1+sourceCell.length()/500);
                String transCell = translationTestPage.getTranslation();
                System.out.println("translatedCell = " + transCell);

                // translation is done as one cell string:


                // problem is below code:
                // how to write translation into each cell?

                //Cell cell = row.createCell(i);
                //cell.setCellValue(translation);
               // row.getCell(i).setCellValue(translation);

               // System.out.println("sheetTarget.getRow(j).getCell(i) = " + sheetTarget.getRow(j).getCell(i));

            }
            System.out.println(" ================== source row(" + j + ") read in ends =============== ");

            FileOutputStream fos = new FileOutputStream(pathTarget);
            workbookW.write(fos);
            fos.close();


            //Driver.closeDriver();
        }


    }




}
