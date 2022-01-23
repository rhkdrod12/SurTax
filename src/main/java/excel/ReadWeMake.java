package excel;

import utils.DataVo;
import utils.Oper;
import utils.Path;
import utils.PricePattern;

import java.io.File;
import java.util.List;

public class ReadWeMake implements ReadExcel {
    
    //읽어올 경로 지정
    String path = Path.WEMAKE_PATH;
    
    @Override
    public DataVo ReadFile() {
    
        String dirPath = prefix + path + suffix;
        
        DataVo weMake = new DataVo("위메프", PricePattern.WeMake);
        
        File dir = new File(dirPath);
        System.out.println(dir.getAbsolutePath());
        if (dir.isDirectory()) {
            File[] files = dir.listFiles();
            
            for (File file : files) {
                
                if (file.isFile() && !file.isHidden()) {
                    String path = file.getPath();
                    
                    List<List<String>> list = ExcelIO.ReadExcel(path, 0, 17, 0, 11);
                    for (List<String> valList : list) {
                        
                        String month = valList.get(0).substring(5, 7);
                        String part = valList.get(4).trim().equals("온라인") ? "기타매출" : "알수없음";
                        String price = valList.get(9);
                        
                        weMake.setPart(month, part, price, Oper.PLUS);
                    }
                }
            }
        } else {
            System.out.println("디렉토리 경로가 올바르지 않습니다.");
        }
        System.out.println(weMake);
        return weMake;
    }
}
