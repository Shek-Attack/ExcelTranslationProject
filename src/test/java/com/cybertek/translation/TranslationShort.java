package com.cybertek.translation;

import com.cybertek.pages.TranslationTestPage;
import com.cybertek.utilities.BrowserUtils;
import com.cybertek.utilities.Driver;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.openqa.selenium.Keys;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;

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
        XSSFWorkbook workbookTarget = new XSSFWorkbook(fsTarget);
        XSSFSheet sheetTarget = workbookTarget.getSheetAt(0);



        // rows:

        String targetWord;
        for (int j = 0; j <6; j++) {
            System.out.print("");

            // go to google translate from "source" to "target":
            Driver.getDriver().get("https://www.google.com/search?q="+source+"+to+"+target[0]+"+translate&oq=chinese+to+english&aqs=chrome.1.69i57j35i39j0i512l8.9361j0j15&sourceid=chrome&ie=UTF-8");

            // columns of a row:
            for (int i = 0; i <colNum; i++) {

                // write into the sourceTextArea

                translationTestPage.sourceTextArea.sendKeys(Keys.CLEAR);
                BrowserUtils.waitForVisibility(translationTestPage.sourceTextArea,30);
                translationTestPage.sourceTextArea.sendKeys(sheetSource.getRow(j).getCell(i)+"|\n", Keys.ENTER);
                System.out.println("sheetSource.getRow("+j+").getCell("+i+") = " + sheetSource.getRow(j).getCell(i));
                // source writing and reading is complete

                //read the translation from targetTextArea
                // wait for translation
                BrowserUtils.waitForVisibility(translationTestPage.targetTextArea, 30);

            }
            System.out.println(" ================== Row("+j+") input is provided as source text =============== ");
            BrowserUtils.waitForVisibility(translationTestPage.targetTextArea, 30);
            targetWord=translationTestPage.targetTextArea.getText();
            System.out.println(targetWord);

           // translation is done as one row string, extract the cell values:
                // problem is below code:

                String[] entryRowCells=targetWord.split("\\|\n");
                String entryRowArray=Arrays.toString(entryRowCells);
            System.out.println("entryRowArray = " + entryRowArray);


            for (int i = 0; i <entryRowCells.length; i++) {
                    XSSFCell cell = sheetTarget.getRow(j).createCell(12);

                    cell.setCellValue(Arrays.toString(entryRowCells));
                    System.out.println("Cell["+i+"] = " + entryRowCells[i]);

                }


            }

        //System.out.println(" ");


        FileOutputStream fos = new FileOutputStream(pathTarget);
        workbookTarget.write(fos);
        fos.close();


        //Driver.closeDriver();


    }




}
