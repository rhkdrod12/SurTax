/**
 *
 */

import excel.*;
import utils.DataVo;

import java.text.ParseException;
import java.util.HashMap;
import java.util.Map;
import java.util.Scanner;

/**
 * @author Kim
 *
 */
public class Main {
    
    
    /***********************
     * @Method main
     * @param args
     * @throws ParseException
     *
     * @COMMENT
     ***********************/
    public static void main(String[] args) throws ParseException {
        
        Map<String, DataVo> map = new HashMap<String, DataVo>();
        
        map.put("배민", new ReadBemin().ReadFile());
        map.put("쿠팡", new ReadCoupang().ReadFile());
        map.put("위메프", new ReadWeMake().ReadFile());
        map.put("요기오", new ReadYugio().ReadFile());
        map.put("카카오", new ReadKakao().ReadFile());
        
        ExcelIO.WriteExcel2(map);
        
        Scanner sc = new Scanner(System.in);
        
        System.out.printf("엔터를 눌러 프로그램 종료");
        sc.nextLine();
        System.out.println("프로그램을 종료합니다.");
        sc.close();
        
    }
}
