package excel;

import utils.DataVo;
import utils.Oper;
import utils.PricePattern;

import java.io.File;
import java.util.List;

public class ReadBemin implements ReadExcel {
    
    //읽어올 경로 지정
    String path = "/배민/매출";
    
    @Override
    public DataVo ReadFile() {
    
        String dirPath = prefix + path;
    
        DataVo beMin = new DataVo("배민", PricePattern.Bemin);
    
        File dir = new File(dirPath);
        System.out.println(dir.getAbsolutePath());
        if (dir.isDirectory()) {
            File[] files = dir.listFiles();
        
            for (File file : files) {
            
                if (file.isFile() && !file.isHidden()) {
                    String path = file.getPath();
                    List<List<String>> list = ExcelIO.ReadExcel(path, 0, 7, 0, 7);
                
                    for (List<String> valList : list) {
                    
                        if(valList != null) {
                            String month = valList.get(0).substring(5, 7);
                            String part = valList.get(2).trim();
                            String price = valList.get(6);
                            beMin.setPart(month, part, price, Oper.PLUS);
                        }
                    }
                }
            }
        }else {
            System.out.println("디렉토리 경로가 올바르지 않습니다.");
        }
        System.out.println(beMin);
        return beMin;
    }
}
