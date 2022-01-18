package utils;

import java.text.DecimalFormat;
import java.text.ParseException;
import java.util.HashMap;
import java.util.Map;

/**
 * 
 */

/**
 * @author Kim
 *
 */
public class PartSales{
	
	protected Map<String, Long> map = new HashMap<>();
	
	protected PricePattern pattern;
	
	public PartSales() {
	}
	
	public PartSales(PricePattern pattern) {
		this.pattern = pattern;
	}
	
	protected String getPattern(){return pattern.getValue();}
	
	public void setVal(String part, String val, Oper oper){
		try {
			Long price = val.equals("") ? 0 : new DecimalFormat(pattern.getValue()).parse(val).longValue();
			if (Oper.PLUS.equals(oper)) {
				map.put(part, map.getOrDefault(part, 0l) + price);
			} else {
				map.put(part, map.getOrDefault(part, 0l) - price);
			}
		} catch (ParseException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	
	public Map<String, Long> getVal() {
		return map;
	}
	
	
	public Long getPartVal(String part) {
		return map.getOrDefault(part, 0l);
	}
	
	@Override
	public String toString() {
		return "PartSale [map=" + map + "]";
	}
	
}
