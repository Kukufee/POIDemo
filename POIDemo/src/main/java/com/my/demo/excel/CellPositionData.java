package com.my.demo.excel;

import java.util.List;
import java.util.Map;

/**
 *
 * @author MaYue
 * @Version CreateTime:2019年4月5日下午10:17:07
 * 用于设置需要填充的数据（填充规则第几行的第几列到第几列）
 */

public class CellPositionData {
	
	private int startColumn;
	
	private int endColumn;
	
	//Key为第几行，value为startColumn到endColumn之间的单元格，从0开始
	//key从0开始
	private Map<Integer,List<String>> data;
	
	/**
	 * 可以同时填充多行
	 * @param startColumn
	 * @param endColumn
	 * @param data
	 */
	public CellPositionData(int startColumn, int endColumn, Map<Integer, List<String>> data) {
		this.startColumn = startColumn;
		this.endColumn = endColumn;
		this.data = data;
	}

	public int getStartColumn() {
		return startColumn;
	}

	public void setStartColumn(int startColumn) {
		this.startColumn = startColumn;
	}

	public int getEndColumn() {
		return endColumn;
	}

	public void setEndColumn(int endColumn) {
		this.endColumn = endColumn;
	}

	public Map<Integer, List<String>> getData() {
		return data;
	}

	public void setData(Map<Integer, List<String>> data) {
		this.data = data;
	}
	
	
}
