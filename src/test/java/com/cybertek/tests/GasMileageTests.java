package com.cybertek.tests;

import com.cybertek.pages.GasMileageCalculatorPage;
import com.cybertek.utilities.Driver;
import io.cucumber.java.hu.De;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;

public class GasMileageTests {

    XSSFWorkbook workbook;
    XSSFSheet sheet;
    FileInputStream fileInputStream;
    FileOutputStream fileOutputStream;
    GasMileageCalculatorPage gasMileageCalculatorPage = new GasMileageCalculatorPage();

    @Test
    public void gas_mileage_calculation() throws IOException {

        Driver.getDriver().get("https://www.calculator.net/gas-mileage-calculator.html");

        // go to US Unit view:
        gasMileageCalculatorPage.usUnits.click();

        //Get the path from Content Root
        String path = "src/test/resources/testData/GasMileageTestData.xlsx";

        //loading the file into inputStream object
        fileInputStream = new FileInputStream(path);

        //loading the workbook to the class
        workbook = new XSSFWorkbook(fileInputStream);

        //opening sheet
        sheet = workbook.getSheet("Sheet1");

        for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {

            XSSFRow currentRow = sheet.getRow(rowNum);

            if (!currentRow.getCell(0).toString().equals("Y")){

                if (currentRow.getCell(6) == null){

                    currentRow.createCell(6);

                }

                currentRow.getCell(6).setCellValue("Skip Requested!");

                continue;

            }


            //===========================================================================
            //entering current odometer reading
            //double current = 123000;

            double current = currentRow.getCell(1).getNumericCellValue();

            //Clearing the input box before entering any data
            gasMileageCalculatorPage.inputCurrentOdo.clear();

            //We need to pass 'double' into input box
            gasMileageCalculatorPage.inputCurrentOdo.sendKeys(String.valueOf(current));

            //===========================================================================
            //Entering previous odometer reading
            //double previous = 122000;
            double previous = currentRow.getCell(2).getNumericCellValue();

            gasMileageCalculatorPage.inputPreviousOdo.clear();
            gasMileageCalculatorPage.inputPreviousOdo.sendKeys(String.valueOf(previous));

            //===========================================================================
            //Entering how many gallons of gas used
            //double gas = 70;
            double gas = currentRow.getCell(3).getNumericCellValue();

            gasMileageCalculatorPage.inputGas.clear();
            gasMileageCalculatorPage.inputGas.sendKeys(String.valueOf(gas));

            //===========================================================================
            //Click to calculate button

            gasMileageCalculatorPage.calculateButton.click();

            //How does the calculation of AVG/MPG work
            // (current-previous)/gallons --> avg MPG

            double expectedResult = (current - previous) / gas;

            System.out.println("expectedResult = " + expectedResult);


            //actual result --> 14.29 mpg --> split " " --> {"14.29", "mpg"}
            String[] actualResultArr = gasMileageCalculatorPage.resultInGas.getText().split(" ");
            System.out.println("actualResultArr = " + actualResultArr[0]);


            String actual = actualResultArr[0];
            String expected = String.valueOf(expectedResult);

            DecimalFormat decimalFormat = new DecimalFormat("#0.00");
            String formattedExpected = decimalFormat.format(expectedResult);

            System.out.println("formattedExpected = " + formattedExpected);

            //===========================================================================
            // PASSING EXPECTED VALUE INTO EXCEL SHEET
            //If the cell is empty, we need to create the cell before being able to pass
            //any data into it
            if (currentRow.getCell(4) == null) {

                currentRow.createCell(4);

            }

            //Passing the 'formattedExpected' into our Excel sheet "Expected" cell
            currentRow.getCell(4).setCellValue(formattedExpected);
            //===========================================================================
            // Passing ACTUAL value into EXCEL SHEET

            if (currentRow.getCell(5) == null) {

                currentRow.createCell(5);

            }

            currentRow.getCell(5).setCellValue(actual);
            //===========================================================================
            //PASSING THE STATUS INTO THE EXCEL FILE

            if (currentRow.getCell(6) == null) {
                currentRow.createCell(6);
            }

            //===========================================================================
            //COMPARING ACTUAL VS EXPECTED AND PASSING THE RESULT INTO EXCEL SHEET

            if (actual.equals(formattedExpected)) {
                //System.out.println("PASS!");
                currentRow.getCell(6).setCellValue("PASS!");
            } else {
                //System.out.println("FAIL!");
                currentRow.getCell(6).setCellValue("FAIL!");
            }

            //===========================================================================
            //ENTERING THE CURRENT TIME WHEN TEST IS RUNNING THAT SPECIFIC LINE

            if (currentRow.getCell(7) == null) {
                currentRow.createCell(7);
            }

            DateTimeFormatter DTF = DateTimeFormatter.ofPattern("hh:mm:ss a");
            currentRow.getCell(7).setCellValue(LocalTime.now().format(DTF));

        }

        //===========================================================================
        //We must write into excel file using .write method, otherwise changes will not be applied

        fileOutputStream = new FileOutputStream(path);
        workbook.write(fileOutputStream);

        workbook.close();
        fileInputStream.close();
        fileOutputStream.close();


    }

}
