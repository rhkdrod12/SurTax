/**
 * 
 */
package utils;

import java.util.HashMap;
import java.util.Map;

/**
 * @author Kim
 *
 */
public class DataVo{
	protected String name;
	protected Map<String, PartSales> map = new HashMap<String, PartSales>();
	protected PricePattern pattern;
	
	public DataVo() {
	}
	
	public DataVo(String name, PricePattern pattern) {
		this.name = name;
		this.pattern = pattern;
	}
	
	public String getName(){
		return name;
	}
	
	public Map<String, PartSales> getData(){
		return map;
	}
	
	public PartSales getPartData(String month){
		return map.get(month);
	}
	
	public void setPart(String month, String part, String price, Oper oper){
		if(map.containsKey(month)) {
			map.get(month).setVal(part, price, oper);
		}else {
			PartSales tempObj = new PartSales(pattern);
			tempObj.setVal(part, price, oper);
			map.put(month, tempObj);
		}
	}
	
	@Override
	public String toString() {
		
		String result = name+"\n";
		
		if(map.isEmpty()) {
			result += "----------------------------------------------------------------\n";
			result += "----------------------------------------------------------------\n";
			return result;
		}
		
		result += "----------------------------------------------------------------\n";
		for(String key : map.keySet()) {
			result += key+"월" + map.get(key)+"\n";
		}
		
		result += "----------------------------------------------------------------\n";
		return result;
	}
	
}
