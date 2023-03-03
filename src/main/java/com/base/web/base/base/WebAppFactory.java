package com.base.web.base.base;

import java.net.URL;

import java.util.Properties;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;

import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class WebAppFactory 
{
    Properties properties = new Properties();
    Properties providerProperties = new Properties();
    Properties seleniumWebProperties = new Properties();
    DesiredCapabilities capabilities = new DesiredCapabilities();
    public static WebDriver driver;
    private static String resourcesPath="/src/main/resources/";
    private static String resourcesPropertiesPath="/src/main/resources/properties/";
    private static String resourcesPropertiesFullPath=System.getProperty("user.dir")+ resourcesPropertiesPath;
    private static String resourcesFullPath=System.getProperty("user.dir")+ resourcesPath;
    public static String seleniumWebPropertiesPath=resourcesPropertiesFullPath+"selenium.web.properties";
    private static String chromeDriverPath=resourcesFullPath + "drivers/chromedriver";
    private static String firefoxDriverPath=resourcesFullPath + "drivers/firefox";
    public static String applicationPropertiesPath=resourcesPropertiesFullPath+"application.properties";
    
    ConfigService configService;
    Tools tools;
    String url;
    String providerName;
    String provider;
    public String webUrl;

    public WebAppFactory()
    {
        configService = new ConfigService();
        tools = new Tools();
    }

    public WebDriver buildDeviceService() throws FileNotFoundException, IOException, Exception
    {
        properties = configService.getProperties(applicationPropertiesPath);
        seleniumWebProperties = configService.getProperties(seleniumWebPropertiesPath);
        providerName = seleniumWebProperties.getProperty("provider");
        webUrl = properties.getProperty("webUrl");
        String browserName = seleniumWebProperties.getProperty("browserName");
        String applicationType = seleniumWebProperties.getProperty("applicationType");
        
        if (applicationType.equalsIgnoreCase("web")){
            if(providerName.equalsIgnoreCase("remote")){
                capabilities.setBrowserName(browserName);
                driver = new RemoteWebDriver(new URL(seleniumWebProperties.getProperty("remoteUrlChrome")), capabilities);
            } else if(providerName.equalsIgnoreCase("local")) {
                if(browserName.equalsIgnoreCase("chrome")) {                    
                    System.setProperty("webdriver.chrome.driver", chromeDriverPath);
                    driver = new ChromeDriver();
                } else if(browserName.equalsIgnoreCase("firefox")) {
                    System.setProperty("webdriver.gecko.driver", firefoxDriverPath);
                    driver = new FirefoxDriver();
                }
                driver.manage().window().maximize();
            }
        } else if (applicationType.equalsIgnoreCase("api")){
            // here code
        }
        return driver;
    }

    public String getWebUrl() throws IOException
    {
        properties = configService.getProperties(applicationPropertiesPath);
        webUrl = properties.getProperty("webUrl");
        return webUrl;
    }
}