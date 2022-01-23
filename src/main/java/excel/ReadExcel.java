package excel;

import utils.DataVo;
import utils.Path;

public interface ReadExcel {
    
    final String prefix = Path.BASE_PATH + "/부가세자료";
    final String suffix = "/매출";
    
    DataVo ReadFile();
}
