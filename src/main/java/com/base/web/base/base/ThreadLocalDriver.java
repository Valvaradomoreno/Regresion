package com.base.web.base.base;

import io.appium.java_client.MobileElement;
import io.appium.java_client.android.AndroidDriver;

import org.openqa.selenium.WebDriver;

public class ThreadLocalDriver {
    private static final ThreadLocal<AndroidDriver<MobileElement>> tlDriver = new ThreadLocal<AndroidDriver<MobileElement>>();
    private static final ThreadLocal<WebDriver> tlWebDriver = new ThreadLocal<WebDriver>();

    public static synchronized void setTLDriver(AndroidDriver<MobileElement> driver) { tlDriver.set(driver); }

    public static synchronized void setTLWebDriver(WebDriver driver) { tlWebDriver.set(driver); }

    public static synchronized AndroidDriver<MobileElement> getTLDriver() {
        return tlDriver.get();
    }

    public static synchronized WebDriver getTLWebDriver() {
        return tlWebDriver.get();
    }
}