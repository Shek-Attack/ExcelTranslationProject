package com.translation.tests;

import com.translation.pages.TranslationTestPage;
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

// SYSTRAN Translate is slower than Google Translate,
// but size range is much higher than Google Translate
// for 300K, 1K row, 12 column, it takes roughly 1 min for 1 row, so 1K/60min=17 hours.
// With SYSTRAN still some field values are missed in translation.
// DeepL has char limit of 5000, as good as SYSTRAN, but extracting translated text is problematic. Good enough for manual translation.

public class ReadTranslateWriteWithSYSTRANXLSX_Nov {

    //WARNING: make sure you have commented and
    // uncommented correct Elements in TranslationTestPage before running the test!!!

    //public String[][] cell2W=new String[rowNums+1][colNums];
    TranslationTestPage translationTestPage=new TranslationTestPage();


    private static final File sourceFile= new File("D:\\xlsx\\Nov15_2021\\uyghurexcelwork\\Karmay.xls"); // read from
    //String sourceFile = "SourceSampleForTesting.xlsx"; // if the file is directly under the project
    private static final File transFile= new File("D:\\xlsx\\Nov15_2021\\uyghurexcelwork\\KarmayEnglish.xls"); // write into


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

        //int rowNums= sheetR.getLastRowNum(); // for real job
        int rowNums=3; //0,1 // for test run
        int colNums=11; // 0,1,2, ...,11

        //reading from cells using for loop:
        //int rowNums= sheetR.getLastRowNum();

        for (int i =0; i <=rowNums; i++) {

            String[] translation=new String[colNums];

            // get rows from the sheet
            HSSFRow rowR=sheetR.getRow(i);
            Row rowW=sheetW.createRow(i);

            for (int j = 0; j < colNums; j++) {
                Driver.getDriver().get("https://translate.systran.net/?source=zh&target=en");

                // get cell from the row
                HSSFCell cellR=rowR.getCell(j);// read/get content from cell(j)
                String cellRContent="";
                // column 5 is Date: we find with experiment that
                // cellR.getRawValue() get the numeric values as numeric and
                // if source cell contains "-", we should keep it as "-": otherwise tow digit numbers show up.
                if(i>0&&(j==5)){
                     cellRContent=(!cellR.toString().contains("-"))? String.valueOf(cellR.getDateCellValue()) :"-";
                }
                else {cellRContent=cellR.toString();}

                if(i>0&&(j==8||j==10)){
                    cellRContent=(cellR.toString().contains("-"))?"-":cellR.toString();
                }
                else {cellRContent=cellR.toString();}

                System.out.println("cellR("+i+","+j+") = " + cellRContent); // see what is there

                //translation begins here==============================

                //read cell from the source file into the sourceTextArea of DeepL

                // source reading is complete
                translationTestPage.sourceTextArea.sendKeys(Keys.CLEAR); // crucial: if missing clear, the first element shows up in the second element and so the last contains everything before.
                translationTestPage.sourceTextArea.sendKeys(cellRContent +"", Keys.ENTER);
                //List<String> sourceList=new ArrayList(Arrays.asList(rowR));
                //System.out.println("sourceList.get("+j+") = " + sourceList.get(j));

                //wait just long enough depending on the length of the last cell.
                int waitTimeForLongChar2Translate=2+(cellRContent.length()/40); //wait chars 2 translate
                //System.out.println("waitTimeForLongChar2Translate = " + waitTimeForLongChar2Translate);
                BrowserUtils.wait(waitTimeForLongChar2Translate);

                //read the translation from targetTextArea
                translation[j]= translationTestPage.transTextArea.getText();
                System.out.println("cellW("+i+","+j+")trans= " + translation[j]);
                String cell2WContent= translation[j];
                //cell2W[i][j]=cell2WContent;
                //System.out.println("cell2W["+i+"]["+j+"] = " + cell2W[i][j]);

                //create a cell to write the translation into:
                Cell cellW= rowW.createCell(j);

                //assign a value to each cell[i,j]:

                // email and phone numbers no need to be translated:
                // only column 5 content need to be kept as original,
                // the rest is String for writing.
                if(i>0&&(j==5|| j==8||j==10)){
                    cellW.setCellValue(cellR.toString());
                }else{
                    cellW.setCellValue(cell2WContent);
                }


            }// End of column j in row i:


            //translation ends here =================================

            System.out.println("============== End of Row "+i+" Read in ===================");
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