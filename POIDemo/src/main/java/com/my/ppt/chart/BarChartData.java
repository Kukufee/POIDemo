package com.my.ppt.chart;

import java.util.Arrays;

/**
 *
 * @author MaYue
 * @Version CreateTime:2019年4月7日下午2:51:34
 *
 */

public class BarChartData implements ChartData{
	
	//图表的位置，获取方式为：
	public int index;
	
	//图表的类型
	public String chartType;
	
	//图表标题
	public String title;
	
	//系列名
	public String[] series;
	
	//横轴
	public String[] categoryName;
	
	//系列数据，每个内层数组代表一个系列的数据，顺序与series顺序一致，对应excel的一列
	public double[][] serieData;

	public BarChartData(String title, String[] series, String[] categoryName, double[][] serieData) {
		this.title = title;
		this.series = series;
		this.categoryName = categoryName;
		this.serieData = serieData;
	}

	public BarChartData(int index, String[] series, String[] categoryName, double[][] serieData) {
		this.index = index;
		this.series = series;
		this.categoryName = categoryName;
		this.serieData = serieData;
	}

	@Override
	public String toString() {
		return "BarChartData [index=" + index + ", chartType=" + chartType + ", title=" + title + ", series="
				+ Arrays.toString(series) + ", categoryName=" + Arrays.toString(categoryName) + ", serieData="
				+ Arrays.toString(serieData) + "]";
	}
	
}
