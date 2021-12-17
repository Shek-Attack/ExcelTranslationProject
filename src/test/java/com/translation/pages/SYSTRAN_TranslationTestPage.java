package com.translation.pages;

import com.translation.utilities.Driver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class SYSTRAN_TranslationTestPage {

    public SYSTRAN_TranslationTestPage(){PageFactory.initElements(Driver.getDriver(), this);}

    //@FindBy(xpath = "(//textarea)[1]") // DeepL
    //@FindBy(xpath = " //textarea") // Google Translate
    @FindBy(xpath = "//textarea") // SYSTRAN text-area
    public WebElement sourceTextArea;

    //@FindBy(xpath = "(//textarea)[2]") // DeepL
    //@FindBy(xpath = "//div[@id='translateContent']") // SYSTRAN translated text-area
    //@FindBy(xpath = " //pre[@id='tw-target-text']") // Google Translate
    //@FindBy(xpath = "//span[@lang='ug']") // Google Translate UYGHUR
    @FindBy(xpath= "//div[@id='outputEditor']") // SYSTRAN text-area
    //@FindBy(xpath = "(//div[@class='systran_output'])[2]") // SYSTRAN text-area
    public WebElement transTextArea;

    public String getTranslation(){
        return transTextArea.getText();
    }



}
