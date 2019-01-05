package top.tangshitai.excel.bean;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;

public class POIResult {
	private String sheetName;
	private int cellLength;
	private List<String[]> data = new ArrayList<String[]>();
	public String getSheetName() {
		return sheetName;
	}
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	
	public POIResult() {
		super();
	}
	public POIResult(String sheetName,int cellLength) {
		super();
		this.sheetName = sheetName;
		this.cellLength = cellLength;
	}
	public void add(String[] rowResult) {
		boolean flag = false;
		for(String t:rowResult) {
			if(!StringUtils.isBlank(t)) flag = true;
		}
		if(flag) data.add(rowResult);
	}
	public String[][] getData() {
		String[][] rowsContentarray = new String[data.size()][cellLength];
		for(int x=0;x<data.size();x++) {
			rowsContentarray[x] = data.get(x);
		}
		return rowsContentarray;
	}
}
