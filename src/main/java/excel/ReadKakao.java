package excel;

import utils.DataVo;
import utils.Oper;
import utils.Path;
import utils.PricePattern;

import java.io.File;
import java.util.List;

public class ReadKakao implements ReadExcel {
    
    //읽어올 경로 지정
    String path = Path.KAKAO_PATH;
    
    @Override
    public DataVo ReadFile() {
    
        String dirPath = prefix + path + suffix;
    
        DataVo kakao = new DataVo("카카오", PricePattern.Kakao);
        
        File dir = new File(dirPath);
        System.out.println(dir.getAbsolutePath());
        if (dir.isDirectory()) {
            File[] files = dir.listFiles();
            
            for (File file : files) {
                if (file.isFile() && !file.isHidden()) {
                    String path = file.getPath();
                    List<List<String>> list = ExcelIO.ReadExcel(path, 0, 1, 0, 7);
    
                    int dateIdx = 0;
                    int partIdx = 3;
                    int priceIdx = 6;
                    
                    for (List<String> valList : list) {
                        
                        String month = valList.get(dateIdx).substring(5, 7);
                        String part = valList.get(partIdx).trim();
                        
                        if (part.contains("카드")) {
                            part = "카드매출";
                        } else if (part.contains("현금성")) {
                            part = "현금매출";
                        } else {
                            part = "기타매출";
                        }
                        
                        String price = valList.get(priceIdx);
                        kakao.setPart(month, part, price, Oper.PLUS);
                    }
                }
            }
        } else {
            System.out.println("디렉토리 경로가 올바르지 않습니다.");
        }
        System.out.println(kakao);
        return kakao;
    }
}
