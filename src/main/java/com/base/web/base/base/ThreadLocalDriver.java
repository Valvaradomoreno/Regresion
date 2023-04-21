package com.base.web.base.base;


import org.openqa.selenium.WebDriver;

public class ThreadLocalDriver {
    private static final ThreadLocal<WebDriver> tlWebDriver = new ThreadLocal<WebDriver>();


    public static synchronized void setTLWebDriver(WebDriver driver) { tlWebDriver.set(driver); }


    public static synchronized WebDriver getTLWebDriver() {
        return tlWebDriver.get();
    }
}