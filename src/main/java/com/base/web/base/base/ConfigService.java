package com.base.web.base.base;

import java.util.Properties;
import java.io.FileReader;
import java.io.IOException;

import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

public class ConfigService 
{
    Properties properties = new Properties();
    JSONObject config;

    public Properties getProperties(String propertiePath) throws IOException
    {
        properties.load(new FileReader(propertiePath));
        return properties;
    }

    public JSONObject getConfigJson(String configPath) throws Exception
    {
        
        JSONParser parser = new JSONParser();
        config = (JSONObject) parser.parse(new FileReader(configPath));
        return config;
    }
}
