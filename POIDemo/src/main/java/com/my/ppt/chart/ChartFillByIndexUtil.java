package com.my.ppt.chart;

import java.util.List;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFSlide;

/**
 *
 * @author MaYue
 * @Version CreateTime:2019年4月8日下午10:16:15
 * 通过下标填写图表
 * 使用该工具首先通过ChartIndexGetUtil获取chart下标
 */

public class ChartFillByIndexUtil {
	
	public static void fillChartByIndex(XSLFSlide slide, List<ChartData> chartDatas) {
		List<POIXMLDocumentPart> parts = slide.getRelations();
		if (parts == null || parts.size() == 0) {
			return;
		}
		
		for (ChartData chartData : chartDatas) {
			
			if (chartData instanceof BarChartData) {
				BarChartData barChartData = (BarChartData) chartData;
				XSLFChart barChart = (XSLFChart)parts.get(barChartData.index);
				BarChartFillUtill.fillBarChart(barChart, barChartData);
			}
			
			if (chartData instanceof PieChartData) {
				PieChartData pieChartData = (PieChartData) chartData;
				XSLFChart pieChart = (XSLFChart)parts.get(pieChartData.index);
				PieChartFillUtil.fillPieChartByData(pieChart, pieChartData);
			}
		}
	}
	
}
