package com.cybertek.translation;

import com.cybertek.pages.TranslationTestPage;
import com.cybertek.utilities.BrowserUtils;
import com.cybertek.utilities.Driver;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.openqa.selenium.Keys;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class TranslationTest {

    XSSFWorkbook workbook;
    XSSFSheet sheet;
    FileInputStream fileInputStream;
    FileOutputStream fileOutputStream;
   TranslationTestPage translationTestPage=new TranslationTestPage();

    // translate from
    String source="chinese"; // can be any language supported by Google Translate
    // to
    String target0="english";
    String target1="german";
    String target2="french";
    String target3="russian";
    String target4="turkish";
    String target5="arabic";
    String target6="japanise";
    String target7="spanish";

    @Test
    public void setTranslationTest(){

        // go to google translate from "source" to "target":
        Driver.getDriver().get("https://www.google.com/search?q="+source+"+to+"+target0+"+translate&oq=chinese+to+english&aqs=chrome.1.69i57j35i39j0i512l8.9361j0j15&sourceid=chrome&ie=UTF-8");

        // write into the sourceTextArea
        String sourceWord="解决方案, 产品中心, 服务支持";
        translationTestPage.sourceTextArea.sendKeys(sourceWord, Keys.ENTER);

        //read the translation from targetTextArea
        // wait 500 milliseconds for translation
        BrowserUtils.sleepMS(5);

        String targetWord =translationTestPage.targetTextArea.getText();
        System.out.println("translation = " + targetWord);


    }


}
