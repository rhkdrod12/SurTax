package excel;

import utils.DataVo;
import utils.Oper;
import utils.Path;
import utils.PricePattern;

import java.io.File;
import java.util.List;

public class ReadNaver implements ReadExcel {
    
    //읽어올 경로 지정
    String path = Path.NAVER_PATH;
    
    @Override
    public DataVo ReadFile() {
    
        String dirPath = prefix + path + suffix;
    
        DataVo naver = new DataVo("네이버", PricePattern.Naver);
        
        File dir = new File(dirPath);
        System.out.println(dir.getAbsolutePath());
        if (dir.isDirectory()) {
            File[] files = dir.listFiles();
            
            for (File file : files) {
                if (file.isFile() && !file.isHidden()) {
                    String path = file.getPath();
                    List<List<String>> list = ExcelIO.ReadExcel(path, 0, 4, 0, 16);
                    for (List<String> valList : list) {
                        String month = valList.get(0).substring(5, 7);
                        String part = "카드매출"; //네이버는 전부 선결재(네이버페이) 방식 -> 카드매출로 지정
                        String price = valList.get(16);
                        naver.setPart(month, part, price, Oper.PLUS);
                    }
                }
            }
        } else {
            System.out.println("디렉토리 경로가 올바르지 않습니다.");
        }
        return naver;
    }
}
