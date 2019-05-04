package com.my.ppt.chart.client;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import com.my.ppt.chart.BarChartData;
import com.my.ppt.chart.ChartData;
import com.my.ppt.chart.ChartFillByIndexUtil;
import com.my.ppt.chart.PieChartData;
import com.my.pptcreatebytemp.PPTCreateUtill;

/**
 *
 * @author MaYue
 * @Version CreateTime:2019年4月8日下午11:42:37
 * 填充第三页图表
 */

public class ClientFillChartByIndex {
	public static void main(String[] args) throws IOException {
		
		XMLSlideShow ppt = PPTCreateUtill.createPPT("\\PPTTempete\\POI测试.pptx");
		//获取所有页数，页数顺序为list元素顺序
		List<XSLFSlide> slides = PPTCreateUtill.getSlides(ppt);
		
		//第一个柱状图数据,测试柱状图
		BarChartData barData1 = createBarData1();
		//第二个柱状图数据,测试柱状图二
		BarChartData barData2 = createBarData2();
		//第一个饼状图数据,测试饼状图
		PieChartData pieData1 = createPieData1();
		
		ArrayList<ChartData> chartDataPage3 = new ArrayList<ChartData>();
		chartDataPage3.add(pieData1);
		chartDataPage3.add(barData1);
		chartDataPage3.add(barData2);
		
		//填充第三页ppt数据
		ChartFillByIndexUtil.fillChartByIndex(slides.get(2), chartDataPage3);
		//生成报告
		FileOutputStream outPutStream = new FileOutputStream(new File("\\PPTTemplete\\测试通过index生成.pptx"));
		ppt.write(outPutStream);
		outPutStream.close();
	}
	
	private static BarChartData createBarData1() {
		String[] series = {"测试系列1", "测试系列2"};
		String[] categoryName = {"测试类型1", "测试类型2", "测试类型3", "测试类型4", "测试类型5"};
		double[][] seriesData = new double[2][5];
		seriesData[0][0] = 0.21;
		seriesData[0][1] = 0.31;
		seriesData[0][2] = 0.41;
		seriesData[0][3] = 0.51;
		seriesData[0][4] = 0.61;
		
		seriesData[1][0] = 0.61;
		seriesData[1][1] = 0.51;
		seriesData[1][2] = 0.41;
		seriesData[1][3] = 0.31;
		seriesData[1][4] = 0.21;
		
		BarChartData barChartData = new BarChartData(3, series, categoryName, seriesData);
		return barChartData;
	}
	
	private static BarChartData createBarData2() {
		String[] series = {"测试系列1"};
		String[] categoryName = {"测试类型1", "测试类型2", "测试类型3"};
		double[][] seriesData = new double[1][3];
		seriesData[0][0] = 0.21;
		seriesData[0][1] = 0.31;
		seriesData[0][2] = 0.41;
		BarChartData barChartData = new BarChartData(2, series, categoryName, seriesData);
		return barChartData;
	}
	
	private static PieChartData createPieData1() {
		String[] series = {"测试系列1"};
		String[] categoryName = {"测试类型1", "测试类型2", "测试类型3", "测试类型4", "测试类型5"};
		double[][] seriesData = new double[1][5];
		seriesData[0][0] = 0.21;
		seriesData[0][1] = 0.31;
		seriesData[0][2] = 0.21;
		seriesData[0][3] = 0.11;
		seriesData[0][4] = 0.15;
		
		PieChartData barChartData = new PieChartData(1, series, categoryName, seriesData);
		return barChartData;
	}
}
