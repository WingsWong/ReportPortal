import ide.watsonwong.general.FileService;
import ide.watsonwong.general.PropertiesService;
import ide.watsonwong.service.ExcelService;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileDescriptor;
import java.io.FileOutputStream;
import java.io.PrintStream;
import java.io.UnsupportedEncodingException;
import java.util.HashMap;
import java.util.List;
import java.util.Properties;
import java.util.Scanner;

public class Application {

    static Properties properties;
    
    static List<List<HashMap>> datas;
    static String sheetName;

    public static void main(String[] args){

        try {
            System.setOut(new PrintStream(new FileOutputStream(FileDescriptor.out), true, "UTF-8"));
        } catch (UnsupportedEncodingException e) {
            throw new InternalError("VM does not support mandatory encoding UTF-8");
        }
        System.out.println(System.getProperty("os.name"));

        properties = new Properties();
        PropertiesService ps = new PropertiesService();
        FileService fs = new FileService();

        try {
            properties = ps.readProperties();
        } catch (Exception e) {
            e.printStackTrace();
        }

        ExcelService excelService = new ExcelService();

        Scanner myObj = new Scanner(System.in);  // Create a Scanner object
        //System.out.println("Enter Sheet Name");
        //String sheetName = myObj.nextLine();
        int before = Integer.parseInt(properties.getProperty("before"));
        int after = Integer.parseInt(properties.getProperty("after"));
        List<Object> firstRow = null;
        resetValue();


        //read excel

        try {
            File[] fList = fs.readFiles(properties.getProperty("importExcel"));

            for(File file: fList) {
                Workbook wbI = excelService.readExcel(file);
                Sheet sheetI = excelService.openSheet(wbI);
                sheetName = sheetI.getSheetName();
                System.out.println("Sheet Name : " + sheetName);
                firstRow = excelService.getFirstRow(sheetI, before, after);
                datas = excelService.getDataRow(sheetI,before,after);

                //write excel
                excelService.writeExcel(firstRow, datas, sheetName,
                properties.getProperty("firstColumn"), properties.getProperty("secondColumn"), before, after,
                properties.getProperty("exportExcel"));
            }


        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    private static void resetValue() {
        datas = null;
        sheetName = null;
    }
}
