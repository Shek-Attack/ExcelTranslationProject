package com.cybertek.pages;

import com.cybertek.utilities.Driver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

public class TranslationTestPage {

    public TranslationTestPage(){PageFactory.initElements(Driver.getDriver(), this);}

    @FindBy(xpath = "//textarea[@id='tw-source-text-ta']")
    public WebElement sourceTextArea;

    @FindBy(xpath = "//pre[@id='tw-target-text']")
    public WebElement targetTextArea;



}
