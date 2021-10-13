package com.cybertek.translation;

import com.cybertek.pages.TranslationTestPage;
import com.cybertek.utilities.BrowserUtils;
import com.cybertek.utilities.Driver;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.openqa.selenium.Keys;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Translation {

   TranslationTestPage translationTestPage=new TranslationTestPage();

    // translate from
    String[] source={"chinese", "russian"}; // can be any other language supported by Google Translate
    // to
    String[] target={"english", "german", "french", "russian", "turkish", "japanese", "arabic", "spanish"};

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

        // rows:

        // title row:

        String targetWord="/";

        for (int j = 0; j <=2; j++) {

            // go to google translate from "source" to "target":
            Driver.getDriver().get("https://www.google.com/search?q="+source[0]+"+to+"+target[0]+"+translate&oq=chinese+to+english&aqs=chrome.1.69i57j35i39j0i512l8.9361j0j15&sourceid=chrome&ie=UTF-8");

            // cells:
            for (int i = 0; i <colNum; i++) {
                Cell cell=sheetSource.getRow(j).getCell(i);

                System.out.print(cell+"/");

                // write into the sourceTextArea
                //translationTestPage.sourceTextArea.sendKeys(Keys.CLEAR);
                translationTestPage.sourceTextArea.sendKeys(" "+cell+"/", Keys.ENTER);

                //read the translation from targetTextArea
                // wait for translation
                //BrowserUtils.waitForVisibility(translationTestPage.targetTextArea, 5);

            }
            System.out.println(" ");
            targetWord=translationTestPage.targetTextArea.getText();
            System.out.println(targetWord);
            System.out.println(" ");


        }



        // write to the Excel file:
        //Path of the Excel file
        String pathTarget = "TargetSample.xlsx";
        FileInputStream fsTarget = new FileInputStream(pathTarget);
        //Creating a workbook
        XSSFWorkbook workbookTarget = new XSSFWorkbook(fsTarget);
        XSSFSheet sheetTarget = workbookTarget.getSheetAt(0);

        int rowNumTarget = sheetTarget.getLastRowNum();
        int colNumTarget = 11; // starting 0,1,2,...,11: so it is 12 columns

        XSSFCell targetCell = sheetTarget.getRow(0).getCell(0);

        System.out.println("Before = " + targetCell);

        targetCell.setCellValue(targetWord);

        System.out.println("After = " + targetCell);


        FileOutputStream fos = new FileOutputStream(pathTarget);
        workbookTarget.write(fos);
        fos.close();


        Driver.closeDriver();


    }




}
