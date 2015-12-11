package utils;

import java.io.*;
import java.util.Properties;

public class PropertiesManager {

    private static PropertiesManager instance = null;
    private static Properties props = null;

    //get the properties file name
    private static String propertiesFilePath="./src/test/config/";
    //private static String userPropertiesFile = propertiesFilePath+"CI.properties";

    private PropertiesManager() {

    }

    public static PropertiesManager getInstance() {
         if (instance == null) {

             instance = new PropertiesManager();
             props = new Properties();

             //make sure something was passed in. If no, set to defaults
             try {
                 String fileName = System.getProperty("SUT");
                 if ((fileName == null) || (fileName.equals(""))) {
                     fileName = "CI";
                 }

                 //load the properties into memory
                 FileInputStream in = new FileInputStream(propertiesFilePath+fileName+".properties");
                 props.load(in);
                 in.close();


             } catch (IOException e) {
                 e.printStackTrace();
             }

             // Dump the properties on debug
             for (String key : props.stringPropertyNames()) {
                 //System.out.println("Property " + key + " = " + props.getProperty(key));
             }
         }
             return instance;
    }

    public String getValue(String _key){
        //System.out.println("Return Property " + _key + " = " + props.getProperty(_key));
        return props.getProperty(_key);
    }


}


//REF: http://java.dzone.com/articles/singleton-design-pattern-%E2%80%93