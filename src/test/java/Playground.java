import ide.watsonwong.general.PropertiesService;
import ide.watsonwong.service.ExcelService;

import java.util.Properties;

public class Playground {

    static Properties properties;

    public static void main(String[] args){

        properties = new Properties();
        PropertiesService ps = new PropertiesService();


        try{
            properties = ps.readProperties();
            String read = properties.getProperty("before");
            System.out.println(read);

        }
        catch (Exception e) {
            e.printStackTrace();
        }
        ExcelService excelService = new ExcelService();

    }
}
