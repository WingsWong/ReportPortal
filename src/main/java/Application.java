import ide.watsonwong.general.PropertiesService;
import ide.watsonwong.service.ExcelService;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

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


    public static void main(String[] args){

        try {
            System.setOut(new PrintStream(new FileOutputStream(FileDescriptor.out), true, "UTF-8"));
        } catch (UnsupportedEncodingException e) {
            throw new InternalError("VM does not support mandatory encoding UTF-8");
        }

        properties = new Properties();
        PropertiesService ps = new PropertiesService();

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
        List<List<HashMap>> datas = null;
        String sheetName = null;

        //read excel

        try {
            Workbook wbI = excelService.readExcel(properties.getProperty("importExcel"));
            Sheet sheetI = excelService.openSheet(wbI);
            sheetName = sheetI.getSheetName();
            System.out.println("Sheet Name : " + sheetName);
            firstRow = excelService.getFirstRow(sheetI, before, after);
            datas = excelService.getDataRow(sheetI,before,after);
//            System.out.println("First row size:" + firstRow.size());
//            int row_size = 2;
//            for(List<HashMap> data: datas) {
//                System.out.println("Data row size:[" + + row_size++ + "]" + data.size());
//                if(data.size() != firstRow.size()) {
//                    System.out.println("Data row size:[" + + row_size++ + "]" + data.size());
//                    for(HashMap mp: data) {
//                        for ( Object key : mp.keySet() ) {
//                            System.out.println(key + "||" + mp.get(key) );
//
//                        }
//                    }
//                }
//            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        //write excel
        excelService.writeExcel(firstRow, datas, sheetName,
                properties.getProperty("firstColumn"), properties.getProperty("secondColumn"), before, after,
                properties.getProperty("exportExcel"));



    }
}
