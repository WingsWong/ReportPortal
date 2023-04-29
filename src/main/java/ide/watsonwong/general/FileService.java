package ide.watsonwong.general;

import java.io.File;
import java.io.FilenameFilter;
import java.util.List;

public class FileService {

    private FilenameFilter getExcelFilter() {
        FilenameFilter filter = new FilenameFilter() {
            @Override
            public boolean accept(File f, String name) {
                // We want to find only .xlsx files
                return name.endsWith(".xlsx");
            }
        };
        return filter;
    }

    
    public File readFile(String path) throws NullPointerException {
        return new File(path);
    }

    public File[] readFiles(String pPath) {
        File path = new File(System.getProperty("user.home") + pPath);

        if(!path.exists()) return null;
        if(!path.isDirectory()) return null;
        
        return path.listFiles();
    }


}
