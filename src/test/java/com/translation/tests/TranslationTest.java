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

public class TranslationTest {

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
        int colNum = 11; // starting 0,1,2,...,10: so it is 11 columns

        System.out.println("rowNum = " + rowNum);
        System.out.println("colNum = " + (colNum+1));

        // rows:

        // title row:

        String targetWord;
        String[] targetRow=new String[12];

        for (int j = 1; j <=2; j++) {
            System.out.print("/");

            // go to google translate from "source" to "target":
            Driver.getDriver().get("https://www.deepl.com/translator#zh/en/");

            // cells:
            for (int i = 0; i <=colNum; i++) {

                System.out.print(sheetSource.getRow(j).getCell(i)+"//");

                // write into the sourceTextArea

                //String sourceWord=sheetSource.getRow(j).getCell(i).getStringCellValue();
                translationTestPage.sourceTextArea.sendKeys(Keys.CLEAR);
                BrowserUtils.waitForVisibility(translationTestPage.sourceTextArea,2);
                translationTestPage.sourceTextArea.sendKeys(sheetSource.getRow(j).getCell(i)+"||", Keys.ENTER);

                //read the translation from targetTextArea
                // wait for translation

            }
            System.out.println("");
            BrowserUtils.wait(2);
            targetWord=translationTestPage.transTextArea.getText();
            System.out.println("targetWord = " + targetWord);

            // System.out.println("targetWord = " + targetWord);
            String word0;
            for (int k = 0; k <=colNum+1; k++) {
                word0=targetWord.substring(0, targetWord.indexOf("||"));

                System.out.println("c"+k+"=" + word0);
                //targetWord=targetWord.replaceFirst(word0+"-",".");
                targetWord=targetWord.replaceFirst(word0+"||","");

            }


        }
        System.out.println(" ");



        // write to the Excel file:
        //Path of the Excel file
        String pathTarget = "TargetSample.xlsx";
        FileInputStream fsTarget = new FileInputStream(pathTarget);
        //Creating a workbook
        XSSFWorkbook workbookTarget = new XSSFWorkbook(fsTarget);
        XSSFSheet sheetTarget = workbookTarget.getSheetAt(0);

        int rowNumTarget = sheetTarget.getLastRowNum();
        int colNumTarget = 11; // starting 0,1,2,...,11: so it is 12 columns

        XSSFCell adamsCell = sheetTarget.getRow(0).getCell(0);

        System.out.println("Before = " + adamsCell);

        adamsCell.setCellValue(targetRow[0]);

        System.out.println("After = " + adamsCell);


        FileOutputStream fos = new FileOutputStream(pathTarget);
        workbookTarget.write(fos);
        fos.close();


        Driver.closeDriver();


    }




}
