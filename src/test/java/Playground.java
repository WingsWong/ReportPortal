import ide.watsonwong.general.FileService;
import ide.watsonwong.general.PropertiesService;
import ide.watsonwong.service.ExcelService;

import java.io.File;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Properties;

public class Playground {

    static Properties properties;

    public static void main(String[] args){

        properties = new Properties();
        PropertiesService ps = new PropertiesService();
        FileService fS = new FileService();

        try{
            properties = ps.readProperties();
            String read = properties.getProperty("importExcel");
            System.out.println(read);

            
            System.out.println(System.getProperty("user.home"));

            File[] fList = fS.readFiles(System.getProperty("user.home"), read);

            for (File f: fList) {
                System.out.println(f.getName());
            }



            

        }
        catch (Exception e) {
            e.printStackTrace();
        }
        ExcelService excelService = new ExcelService();



    }
}
