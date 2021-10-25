package com.cybertek.translation;

import com.cybertek.pages.TranslationTestPage;
import com.cybertek.utilities.BrowserUtils;
import com.cybertek.utilities.Driver;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.openqa.selenium.Keys;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ReadTranslateXLSX {

    int rowNums=1; //0,1
    int colNums=12; // 0,1,2, ...,11

    public String[][] cell2W=new String[rowNums+1][colNums];
    TranslationTestPage translationTestPage=new TranslationTestPage();

    private static final File sourceFile= new File("D:\\xlsx\\SourceSample.xlsx");
    //String sourceFile = "SourceSample.xlsx"; // if the file is directly under the project
    private static final File transFile= new File("D:\\xlsx\\Trans.xlsx");

    // translate from
    String source="chinese"; // can be any language supported by Google Translate
    // to
    String[] target={"english", "german", "french", "russian", "turkish", "japanise", "arabic", "spanish"};


    @Test
    public void setTranslationTest() throws IOException {

        // For reading, need to define source for FileInputStream, workbook, sheet
        FileInputStream fisReading = new FileInputStream(sourceFile);
        XSSFWorkbook workbookR = new XSSFWorkbook (fisReading);
        XSSFSheet sheetR = workbookR.getSheetAt(0);

        // For writing, need to define source for FileInputStream, workbook, sheet
        FileInputStream fisWriting = new FileInputStream(transFile);
        Workbook workbookW = new XSSFWorkbook (fisWriting);
        Sheet sheetW = workbookW.getSheetAt(0);

        //reading from cells using for loop:
        //int rowNums= sheetR.getLastRowNum();
        int rowNums=1;
        int colNums=12;

        for (int i = 0; i <=rowNums; i++) {

            String[] translation=new String[12];

            // get rows from the sheet
            XSSFRow rowR=sheetR.getRow(i);
            Row rowW=sheetW.getRow(i);

            //create colNums of cells in the i row rowW for writing:

            // go to DeppL to translate from "source" to "target":
            Driver.getDriver().get("https://www.google.com/search?q=" + source + "+to+" + target[0] + "+translate&oq=chinese+to+english&aqs=chrome.1.69i57j35i39j0i512l8.9361j0j15&sourceid=chrome&ie=UTF-8");

            //Driver.getDriver().get("https://www.deepl.com/translator#zh/en/");



            for (int j = 0; j < colNums; j++) {


                // get cell from the row
                XSSFCell cellR=rowR.getCell(j);// read/get content from cell(j)
                String cellRContent=cellR.toString();
                System.out.println("cellR("+i+","+j+") = " + cellRContent); // see what is there

                //translation begins here==============================

                //read cell from the source file into the sourceTextArea of DeepL

                // source reading is complete
                translationTestPage.sourceTextArea.clear(); // crutial: if missing clear, the first element shows up in the second element and so the last contains everything before.
                translationTestPage.sourceTextArea.sendKeys(cellRContent +"", Keys.ENTER);
                //List<String> sourceList=new ArrayList(Arrays.asList(rowR));
                //System.out.println("sourceList.get("+j+") = " + sourceList.get(j));

                //wait just long enough depending on the length of the last cell.
                int waitTimeForLongChar2Translate=1+(cellRContent.length()/1000); //wait 1 sec for 1k chars 2 translate
                //System.out.println("waitTimeForLongChar2Translate = " + waitTimeForLongChar2Translate);
                BrowserUtils.wait(waitTimeForLongChar2Translate);

                //read the translation from targetTextArea
                translation[j]= translationTestPage.transTextArea.getText();
                System.out.println("cellW("+i+","+j+")trans= " + translation[j]);
                String cell2WContent= translation[j];
                cell2W[i][j]=cell2WContent;
                System.out.println("cell2W[i][j] = " + cell2W[i][j]);

                //create a cell to write the translation into:


                //assign a value to each cell[i,j]:
                //cellW.setCellValue("row"+i+"cell"+j);
                //cellW.setCellValue(cell2WContent);


            }// End of column j in row i:


            //translation ends here =================================

            System.out.println("============== End of Row "+i+" Read in ===================");

            fisWriting.close();
            FileOutputStream fos =new FileOutputStream(transFile);
            workbookW.write(fos);
            fos.close();
            System.out.println("Done: values are written in "+transFile);


        }// End of row , end of Excel Table

        fisReading.close();

        Driver.closeDriver();


    } // End of Test


} // end of class
