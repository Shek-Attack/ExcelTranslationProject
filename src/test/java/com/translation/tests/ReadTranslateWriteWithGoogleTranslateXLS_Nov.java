package com.translation.tests;

import com.translation.pages.Google_TranslationTestPage;
import com.translation.utilities.BrowserUtils;
import com.translation.utilities.Driver;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import org.junit.Test;
import org.openqa.selenium.Keys;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;


// SYSTRAN Translate is slower than Google Translate,
// but size range is much higher than Google Translate
// for 300K, 1K row, 12 column, it takes roughly 1 min for 1 row, so 1K/60min=17 hours.
// With SYSTRAN still some field values are missed in translation.
// DeepL has char limit of 5000, as good as SYSTRAN, but extracting translated text is problematic. Good enough for manual translation.
// Google Translate is 4 times faster than SYSTRAN Translate.

//Translates xls files, not xlsx files!!!

public class ReadTranslateWriteWithGoogleTranslateXLS_Nov {

    //WARNING: make sure you have commented and
    // uncommented correct Elements in TranslationTestPage before running the test!!!
    // check also colNums. It has to be correct.

    Google_TranslationTestPage translationTestPage=new Google_TranslationTestPage();

    private static final File sourceFile= new File("D:\\xlsx\\Nov15_2021\\ChineseDocs\\Altay.xls"); // read from
    //String sourceFile = "SourceSampleForTesting.xlsx"; // if the file is directly under the project
    private static final File transFile= new File("D:\\xlsx\\Nov15_2021\\Trans4.xls"); // write into


    // translate from
    String source="chinese"; // can be any language supported by Google Translate
    // translate into
    String[] target={"english", "german", "french", "russian", "turkish", "japanese", "arabic", "uyghur"};


    @Test
    public void translationTest() throws IOException {

        // For reading, need to define source for FileInputStream, workbook, sheet
        FileInputStream fisReading = new FileInputStream(sourceFile);
        HSSFWorkbook workbookR = new HSSFWorkbook (fisReading);
        HSSFSheet sheetR = workbookR.getSheetAt(0);

        // For writing, need to define source for FileInputStream, workbook, sheet
        FileInputStream fisWriting = new FileInputStream(transFile);
        Workbook workbookW = new HSSFWorkbook (fisWriting);
        Sheet sheetW = workbookW.getSheetAt(0);


        // row numbers of the sheet:
        //int rowNums= sheetR.getLastRowNum(); // for real job
        int rowNums=3; //0,1, ... // for testing only
        System.out.println("rowNums = " + rowNums);

        // column numbers of the sheet:
        int colNums=sheetR.getRow(0).getLastCellNum(); // 0,1,2, ...,9 // sheet colNums=10 :::
        System.out.println("colNums = " + colNums);

        // if the code does not run after one row, check the column number in the xls file

        for (int i =0; i <=rowNums; i++) {
            String[] translation=new String[colNums];
            // get rows from the sheet
            HSSFRow rowR=sheetR.getRow(i);
            Row rowW=sheetW.createRow(i);

            // go to DeppL/Google  to translate from "source" to "target":
           Driver.getDriver().get("https://www.google.com/search?q=" + source + "+to+" + target[4] + "+translate&oq=chinese+to+english&aqs=chrome.1.69i57j35i39j0i512l8.9361j0j15&sourceid=chrome&ie=UTF-8");
           //Driver.getDriver().get("https://translate.google.ca/?sl=auto&tl=ug");


            for (int j = 0; j < colNums; j++) {

                    // get cell from the row
                    HSSFCell cellR = rowR.getCell(j);// read/get content from cell(j)

                    String cellRContent = cellR.toString();

                    System.out.println("cellR(" + i + "," + j + ") = " + cellRContent); // see what is there

                    if (cellRContent.length() >= 3900) { // Google limit: upto 3900 chars
                        System.out.println("=============Cell [" + i + ", " + j + "] has too many chars=================");
                        //cellRContent=cellRContent.substring(0, 3900);
                        cellRContent = "!!!";
                    }
                    // source reading is complete

                    //translation begins here==============================

                    translationTestPage.sourceTextArea.clear(); // crucial: if missing clear, the first element shows up in the second element and so the last contains everything before.
                    translationTestPage.sourceTextArea.sendKeys(cellRContent + "", Keys.ENTER);

                    //wait just long enough depending on the length of the last cell.
                    int waitTimeForLongChar2Translate = 1 + (cellRContent.length() / 1000); //wait 1 sec for 1k chars 2 translate
                    //System.out.println("waitTimeForLongChar2Translate = " + waitTimeForLongChar2Translate);
                    BrowserUtils.wait(waitTimeForLongChar2Translate);

                    //read the translation from targetTextArea
                    translation[j] = translationTestPage.transTextArea.getText();
                    System.out.println("cellW(" + i + "," + j + ")trans= " + translation[j]);
                    String cell2WContent = translation[j];
                    //cell2W[i][j]=cell2WContent;
                    //System.out.println("cell2W[i][j] = " + cell2W[i][j]);

                    //create a cell to write the translation into:
                    Cell cellW = rowW.createCell(j);
                    //assign a value to each cell[i,j]:
                    //cellW.setCellValue("row"+i+"cell"+j);

                    // Date, email and phone numbers no need to be translated: 5, 8, 10 : do nothing!
                    if(i>0 && (j==5||j==8||j==10)) {
                        cellW.setCellValue(cellR.toString());
                    } else {
                        cellW.setCellValue(cell2WContent);
                    }

            }// End of column j in row i:

            //translation ends here =================================

            System.out.println("============== End of Row "+i+" ===================");
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
