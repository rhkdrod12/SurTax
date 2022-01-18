package excel;

import utils.DataVo;
import utils.Oper;
import utils.PricePattern;

import java.io.File;
import java.util.List;

public class ReadYugio implements ReadExcel {
    
    //읽어올 경로 지정
    String path = "/요기오/매출";
    
    @Override
    public DataVo ReadFile() {
        
        String dirPath = prefix + path;
    
        DataVo yugio = new DataVo("요기오", PricePattern.Yugio);
        
        File dir = new File(dirPath);
        System.out.println(dir.getAbsolutePath());
        if (dir.isDirectory()) {
            File[] files = dir.listFiles();
            
            for (File file : files) {
                if (file.isFile() && !file.isHidden()) {
                    String path = file.getPath();
                    List<List<String>> list = ExcelIO.ReadExcel(path, 0, 20, 0, 10);
                    
                    for (List<String> valList : list) {
                        
                        String month = valList.get(0).substring(5, 7);
                        String part = valList.get(7).trim();
                        
                        if (part.contains("신용카드")) {
                            part = "카드매출";
                        } else if (part.contains("현금")) {
                            part = "현금매출";
                        } else {
                            part = "기타매출";
                        }
                        
                        String price = valList.get(8);
                        yugio.setPart(month, part, price, Oper.PLUS);
                    }
                }
            }
        } else {
            System.out.println("디렉토리 경로가 올바르지 않습니다.");
        }
        System.out.println(yugio);
        return yugio;
    }
}
