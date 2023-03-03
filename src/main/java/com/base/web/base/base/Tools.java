package com.base.web.base.base;

import java.text.SimpleDateFormat;
import java.util.Date;

import java.io.*;

public class Tools 
{
    public String getDateNow()
    {
        SimpleDateFormat formatter = new SimpleDateFormat("ddMMYYYYHHmmss");
        Date date = new Date(System.currentTimeMillis());        
        return formatter.format(date);
    }
}
