package com.thoughtworks;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class Config
{
    private static Config instance = null;

    private static Properties properties = new Properties();

    private Config()
    {
        System.out.println("Entering Config()");

        try
        {
            InputStream inputStream = getClass().getClassLoader().getResourceAsStream("config.properties");
            properties.load(inputStream);
        }
        catch (IOException ex)
        {
            ex.printStackTrace();
        }
    }

    public static void include(String[] args)
    {
        if (instance == null) instance = new Config();

    }

    public static String valueOf(String propertName)
    {
        if (instance == null) instance = new Config();
        return properties.getProperty(propertName);

    }
}
