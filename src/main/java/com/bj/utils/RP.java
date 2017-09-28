package com.bj.utils;


import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

/**
 * Created by JACKSON on 2017/8/26.
 */

public class RP {

    public static String getProperty(String key) throws IOException {
        Properties props = new Properties();
        try {
            InputStream is = RP.class.getResourceAsStream("/config.properties");
            props.load(is);
        }catch (Exception e){
            e.printStackTrace();
        }
        return props.getProperty(key);
    }

    public static void main(String[] args) throws IOException {
        System.out.println(RP.getProperty("Backend_URL"));
        System.out.println(RP.getProperty("Admin"));
        System.out.println(RP.getProperty("Password"));
        System.out.println(RP.getProperty("PeriodPassword"));
    }
}
