package ide.watsonwong.general;

import java.io.*;
import java.util.Properties;

public class PropertiesService {


    private InputStream fileReader() throws Exception{
        try {
            System.out.println(this.getClass().getName());

            InputStream inputStream = this.getClass().getResourceAsStream("/config.properties");
            return inputStream;
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        }
    }

    public Properties readProperties() throws Exception {
        final Properties properties = new Properties();
        properties.load(fileReader());
        return properties;
    }


}
