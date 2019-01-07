package top.tangshitai.excel.bean;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;

public class POIResult {
	private String sheetName;
	private int cellLength;
	private String[] headParams;
	private List<String[]> data = new ArrayList<String[]>();

	public String getSheetName() {
		return sheetName;
	}

	public POIResult() {
		super();
	}

	public POIResult(String sheetName, int cellLength, String[] headParams) {
		super();
		this.sheetName = sheetName;
		this.cellLength = cellLength;
		this.headParams = headParams;
	}

	public void add(String[] rowResult) {
		boolean flag = false;
		for (String t : rowResult) {
			if (!StringUtils.isBlank(t)) flag = true;
		}
		if (flag) data.add(rowResult);
	}

	/**
	 * 返回行的长度
	 * @return
	 */
	public int rowLength() {
		return data.size();
	}

	/**
	 * 获取某行记录中的某个字段的值，根据标题头名称获取
	 * @param rowIndex
	 * @param head
	 * @return
	 */
	public String getRowOfField(int rowIndex, String head) {
		int index = -1;
		for (int x = 0; x < headParams.length; x++) {
			if (headParams[x].equals(head))
				index = x;
		}
		return data.get(rowIndex)[index];
	}

	/**
	 * 获取某行记录中的某个字段的值，根据字段的索引获取
	 * @param rowIndex
	 * @param fieldIndex
	 * @return
	 */
	public String getRowOfField(int rowIndex, int fieldIndex) {
		return data.get(rowIndex)[fieldIndex];
	}

	/**
	 * 获取某一行的所有字段值
	 * @param rowIndex
	 * @return
	 */
	public String[] getRow(int rowIndex) {
		return data.get(rowIndex);

	}
}