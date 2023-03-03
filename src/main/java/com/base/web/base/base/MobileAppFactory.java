package com.base.web.base.base;

import java.net.URL;

import java.util.Map;
import java.util.Iterator;
import java.util.Properties;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.json.simple.JSONObject;

import org.openqa.selenium.remote.DesiredCapabilities;

import io.appium.java_client.android.AndroidDriver;

import io.appium.java_client.MobileElement;

public class MobileAppFactory 
{
    Properties properties = new Properties();
    Properties providerProperties = new Properties();
    DesiredCapabilities capabilities = new DesiredCapabilities();    
    AndroidDriver<MobileElement> driver;
    private static String resourcesPropertiesPath="/src/main/resources/properties/";
    private static String resourcesConfigPath="src/main/resources/config/";
    private static String resourcesPropertiesFullPath=System.getProperty("user.dir")+ resourcesPropertiesPath;
    private static String propertiePath=resourcesPropertiesFullPath +"selenium.mobile.properties";
    private static String androidDevicesconfigFile=".android.devices.config.json";
    private static String appMobileConfigFile=".appmobile.config.json";
    
    ConfigService configService;
    Tools tools;
    String url;
    JSONObject androidDevicesconfig;
    JSONObject appConfig;
    JSONObject appMobileConfig;
    String providerName;
    String provider;
    public String appPackage;
    String device = "";

    public MobileAppFactory()
    {
        configService = new ConfigService();
        tools = new Tools();
    }

    public AndroidDriver<MobileElement> buildDeviceService() throws FileNotFoundException, IOException, Exception
    {
        properties = configService.getProperties(propertiePath);
        providerName = properties.getProperty("providerName");
        
        if(properties.getProperty("mobileType").toString().equalsIgnoreCase("android")){
            device = getDevice().equalsIgnoreCase("") ? properties.getProperty("devices") : getDevice();
            provider = properties.getProperty("provider");
            String app = properties.getProperty("app");
            String appConfigPath = resourcesConfigPath + app + ".config.json";
            appConfig = configService.getConfigJson(appConfigPath);
            appPackage = appConfig.get("appPackage").toString();
            if(provider.equalsIgnoreCase("cloud")){
                providerProperties = configService.getProperties(resourcesPropertiesFullPath + providerName + ".properties");

                String server = providerProperties.getProperty("server");
                if(server == null) {
                    System.out.println("There is not exist a server for cloud");
                }                
                
                String username = providerProperties.getProperty("username");
                if(username == null) {
                    System.out.println("There is not exist a username for cloud");
                }

                String accessKey = providerProperties.getProperty("access_key");
                if(accessKey == null) {
                    System.out.println("There is not exist a accessKey for cloud");
                }

                androidDevicesconfig = configService.getConfigJson(resourcesConfigPath + providerName + androidDevicesconfigFile);
                appMobileConfig = configService.getConfigJson(resourcesConfigPath + providerName + appMobileConfigFile);
                if(providerName.equalsIgnoreCase("lambdatest")){
                    url = "https://"+username+":"+accessKey+"@"+server+"/wd/hub";
                } else if(providerName.equalsIgnoreCase("browserstack")) {
                    capabilities.setCapability("browserstack.user", username);
                    capabilities.setCapability("browserstack.key", accessKey);
                    url = "https://"+server+"/wd/hub";
                }
            } else if(provider.equalsIgnoreCase("local")) {                
                androidDevicesconfig = configService.getConfigJson(resourcesConfigPath + provider + androidDevicesconfigFile);
                appConfig.remove("app");
                appMobileConfig = configService.getConfigJson(resourcesConfigPath + provider + appMobileConfigFile);
                JSONObject devices = (JSONObject) androidDevicesconfig.get("environments");
                JSONObject deviceLocal = (JSONObject) devices.get(device);
                String port = deviceLocal.get("port").toString();                
                url = String.format("http://127.0.0.1:%s/wd/hub", port);
                deviceLocal.remove("port");
                devices.put(device, deviceLocal);
                androidDevicesconfig.put("environments", devices);
            }
        }
        capabilities = getCapabilities((JSONObject) androidDevicesconfig.get("environments"),
                                (JSONObject) appConfig, appMobileConfig);
        driver = new AndroidDriver<MobileElement>(new URL(url), capabilities);
        return driver;
    }

    DesiredCapabilities getCapabilities(JSONObject devicesConfig, JSONObject appConfig, 
        JSONObject appMobileConfig) throws FileNotFoundException, IOException, Exception
    {
        System.out.println(device);
        capabilities = insertCapabilities(capabilities, (JSONObject) devicesConfig.get(device));            
        capabilities = insertCapabilities(capabilities, (JSONObject) appConfig);
        if(provider.equalsIgnoreCase("cloud")){
            JSONObject apps = (JSONObject) appConfig.get("app");
            capabilities.setCapability("app", apps.get(providerName));
        }

        capabilities = insertCapabilities(capabilities, (JSONObject) appMobileConfig);
        String buildId = tools.getDateNow();
        capabilities.setCapability("build", appMobileConfig.get("build") + buildId);
        System.out.println("build: " + appMobileConfig.get("build") + buildId);
        return capabilities;
    }

    public DesiredCapabilities insertCapabilities(DesiredCapabilities capabilities, JSONObject config) 
    {
        Map<String, String> commonCapabilities = (Map<String, String>) config;
        Iterator it = commonCapabilities.entrySet().iterator();
        while (it.hasNext()) {
            Map.Entry pair = (Map.Entry)it.next();
            if(capabilities.getCapability(pair.getKey().toString()) == null){
                capabilities.setCapability(pair.getKey().toString(), pair.getValue().toString());
            }
        }
        return capabilities;
    }

    public String getDevice()
    {
        return this.device;
    }

    public String setDevice(String device)
    {
        return this.device = device;
    }
}
