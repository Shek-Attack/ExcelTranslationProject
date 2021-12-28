package com.translation.tests;

import com.translation.pages.Google_TranslationTestPage;
import com.translation.pages.SYSTRAN_TranslationTestPage;
import com.translation.utilities.BrowserUtils;
import com.translation.utilities.Driver;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Keys;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Option I: use all 3 translators
 * The program should be able to pick the right translation service
 * according to the size of the cell: below 430 Chinese words (3900 English chars),
 * Google Translation (small, but fast); else if below 2000 Chinese words,
 * SYSTRAN Translation(upto 2000 Chinese words and accurate, but 4 times slower than Google);
 * else if below 5000 Chinese words, DeepL Translation(upto 5000 Chinese words);
 * else (no known free service can do it), either keep the cell blank and then do it manually
 * or divide the cell into groups of 5000 chinese words, translate on DeepL and then
 * merge them into one String and write into the cell.
 *
 * Option II:  doing everything with Google Translate:
 * divide the cell which has more than 400 Chinese words;
 * translate each small part and then combine them into one before writing into translation.
 * Ex: 3570 Chinese words:  1+(3570/400)=1+8=9, divided into 9 sub cells and then merged into one.
 *
 * Option III: due to its great quality, it is better go ahead with SYSTRANS.
 *
 * Since Google is not that great in terms of quality, the Option II is not followed.
 */
public class NoEmailNoPhoneGoogleSYSTRANSTranslationXLSX {

    //source and translation files:
    public static final File sourceFile= new File("D:\\xlsx\\Dec_27_2021\\XinjiangPrisonsPage2.xlsx"); // read from
    //static String sourceFile = "SourceSampleForTesting.xlsx"; // if the file is directly under the project
    private static final File transFile= new File("D:\\xlsx\\Dec_27_2021\\XinjiangPrisonsPage2Translation.xlsx"); // to write into. A blank file with the name "Trans6.xlsx" must exist in the folder or under the project.
    //static String transFile= "Trans6.xlsx"; // if the file is directly under the project // to write into. A blank file with the name "Trans6.xlsx" must exist in the folder or under the project.

    // translate from
    static String source="chinese"; // can be any language supported by Google Translate
    // translate into
    static String[] target={"english", "german", "french", "russian", "turkish", "japanese", "arabic", "uyghur"};

    static Google_TranslationTestPage googleTranslationTestPage=new Google_TranslationTestPage();
    static SYSTRAN_TranslationTestPage systran_translationTestPage=new SYSTRAN_TranslationTestPage();

    public static void main(String[] args) throws IOException {

        System.out.println("Source File = " + sourceFile);

        // For reading, need to define source for FileInputStream, workbook, sheet
        FileInputStream fisReading = new FileInputStream(sourceFile);
        XSSFWorkbook workbookR = new XSSFWorkbook(fisReading);
        XSSFSheet sheetR = workbookR.getSheetAt(0);

        // For writing, need to define source for FileInputStream, workbook, sheet
        FileInputStream fisWriting = new FileInputStream(transFile);
        Workbook workbookW = new XSSFWorkbook (fisWriting);
        Sheet sheetW = workbookW.getSheetAt(0);

        // row numbers of the sheet:
        int rowNums= sheetR.getLastRowNum(); // for real job
        //int rowNums=2; //0,1, ... // for testing only
        System.out.println("rowNums = " + rowNums);

        // column numbers of the sheet:
        int colNums=sheetR.getRow(0).getLastCellNum(); // 0,1,2, ...,9 // sheet colNums=10 :::
        System.out.println("colNums = " + colNums);

        // if the code does not run after one ROW, check the column number in the xls file
        for (int i=0; i <rowNums; i++) {
            String[] translation = new String[colNums];
            // get rows from the sheet
            XSSFRow rowR = sheetR.getRow(i); // for reading, just get the row
            Row rowW = sheetW.createRow(i);  // for writing, need to create the row

            for (int j = 0; j < colNums; j++) {

                // get cell from the row
                XSSFCell cellR = rowR.getCell(j);// read/get content from cell(j)

                /**
                 * column 5 reading should be different!!!!!!!!!!!. Treated differently.
                 */

                String cellRContent = cellR.toString();
                int cellLength = cellRContent.length();
                System.out.println("cellR(" + i + "," + j + ") = " + cellRContent); // see what is there
                //System.out.println("cellR(" + i + "," + j + ") :rowR.getCell(j) = " + rowR.getCell(j)); // same to the above line
                // source reading is complete

                int waitTimeForLongChar2Translate = 0;

                if (cellLength < 425) { // Google limit: upto 428 Chinese Words

                    // go to Google  to translate from "source" to "target":
                    Driver.getDriver().get("https://www.google.com/search?q=" + source + "+to+" + target[0] + "+translate&oq=chinese+to+english&aqs=chrome.1.69i57j35i39j0i512l8.9361j0j15&sourceid=chrome&ie=UTF-8");
                    //translation begins here==============================
                    googleTranslationTestPage.sourceTextArea.clear(); // crucial: if missing clear, the first element shows up in the second element and so the last contains everything before.
                    googleTranslationTestPage.sourceTextArea.sendKeys(cellRContent + "", Keys.ENTER);

                    //wait just long enough depending on the length of the last cell.
                    waitTimeForLongChar2Translate = 1 + cellLength / 400; //wait 1 sec for 400 chars 2 translate
                    //System.out.println("waitTimeForLongChar2Translate = " + waitTimeForLongChar2Translate);
                    BrowserUtils.wait(waitTimeForLongChar2Translate);

                    //read the translation from targetTextArea
                    translation[j] = googleTranslationTestPage.getTranslation();

                } else {
                    // go to DeepL translate
                    Driver.getDriver().get("https://translate.systran.net/?source=zh&target=en");

                    System.out.println("=============Cell [" + i + ", " + j + "] has " + cellLength + " Chinese chars=================");
                    //translation begins here==============================
                    systran_translationTestPage.sourceTextArea.clear(); // crucial: if missing clear(), the first element shows up in the second element and so the last contains everything before.
                    systran_translationTestPage.sourceTextArea.sendKeys(cellRContent + "", Keys.ENTER);

                    //wait just long enough depending on the length of the last cell.
                    waitTimeForLongChar2Translate = 6 + (cellLength/200);
                    //wait initially 5 sec and 1 sec for 200 chars 2 translate.
                    // For 800-1000 Chinese chars, at least 10 sec is needed.
                    //wait until the text is written into the source text area
                    BrowserUtils.waitForVisibility(systran_translationTestPage.sourceTextArea, waitTimeForLongChar2Translate);

                    //wait for the translation to show up in the output TextArea
                    BrowserUtils.wait(waitTimeForLongChar2Translate);
                    //read the translation from the output TextArea
                    translation[j] = systran_translationTestPage.getTranslation();

                }

                String cell2WContent = translation[j];

                //create a cell to write the translation into:
                Cell cellW = rowW.createCell(j);


                cellW.setCellValue(cell2WContent);


                System.out.println("cellW(" + i + "," + j + ") = " + cellW);

            }// End of column j in row i:

            //translation ends here =================================

            System.out.println("============== End of Row " + i + " ===================");
        }// End of row , end of Excel Table

        fisWriting.close();
        FileOutputStream fos =new FileOutputStream(transFile);
        workbookW.write(fos);
        fos.close();
        System.out.println("Done: values are written in "+transFile);

        fisReading.close();
        Driver.closeDriver();

    } // End of Test

} // end of class

