package excel;

import utils.DataVo;
import utils.Oper;
import utils.Path;
import utils.PricePattern;

import java.io.File;
import java.util.List;

public class ReadCoupang implements ReadExcel {
    
    //읽어올 경로 지정
    String path = Path.COUPANG_PATH;
    
    @Override
    public DataVo ReadFile() {
    
        String dirPath = prefix + path + suffix;
    
        File dir = new File(dirPath);
    
        DataVo coupang = new DataVo("쿠팡", PricePattern.Coupang);
        
        System.out.println(dir.getAbsolutePath());
        if (dir.isDirectory()) {
            File[] files = dir.listFiles();
            for (File file : files) {
                if (file.isFile() && !file.isHidden()) {
                    String path = file.getPath();
                    String month = path.substring(path.lastIndexOf(".xlsx") - 2, path.lastIndexOf(".xlsx"));
                
                    List<List<String>> list = ExcelIO.ReadExcel(path, 0, 1, 2, 12);
                    for (List<String> valList : list) {
                    
                        String cardSale = valList.get(1);
                        String cardRoll = valList.get(5);
                        coupang.setPart(month, "카드매출", cardSale, Oper.PLUS);
                        coupang.setPart(month, "카드매출", cardRoll, Oper.MINUS);
                    
                        String cashSale = valList.get(2);
                        String cashRoll = valList.get(6);
                        coupang.setPart(month, "현금매출", cashSale, Oper.PLUS);
                        coupang.setPart(month, "현금매출", cashRoll, Oper.MINUS);
                    
                        String otherSale = valList.get(3);
                        String otherRoll = valList.get(7);
                        coupang.setPart(month, "기타매출", otherSale, Oper.PLUS);
                        coupang.setPart(month, "기타매출", otherRoll, Oper.MINUS);
                    }
                }
            }
        }else {
            System.out.println("디렉토리 경로가 올바르지 않습니다.");
        }
        System.out.println(coupang);
        return coupang;
    }
}
