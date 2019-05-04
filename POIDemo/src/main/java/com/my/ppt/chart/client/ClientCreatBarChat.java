package com.my.ppt.chart.client;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import com.my.ppt.chart.BarChartData;
import com.my.ppt.chart.BarChartFillUtill;
import com.my.pptcreatebytemp.PPTCreateUtill;

/**
 *
 * @author MaYue
 * @Version CreateTime:2019年4月7日下午4:41:25
 * 通过图表名填充
 */

public class ClientCreatBarChat {
	public static void main(String[] args) throws IOException {
		XMLSlideShow ppt = PPTCreateUtill.createPPT("E:\\POI测试用PPT\\POI测试.pptx");
		//获取所有页数，页数顺序为list元素顺序
		List<XSLFSlide> slides = PPTCreateUtill.getSlides(ppt);
		
		String[] series = {"测试系列1", "测试系列2"};
		//String[] series = {"测试系列1", "测试系列2", "测试系列3"};
		//String[] series = {"测试系列1"};
		String[] categoryName = {"测试类型1", "测试类型2", "测试类型3", "测试类型4", "测试类型5"};
		//String[] categoryName = {"测试类型1", "测试类型2", "测试类型3", "测试类型4"};
		//String[] categoryName = {"测试类型1", "测试类型2", "测试类型3"};
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
		
		BarChartData barChartData = new BarChartData("测试柱状图2", series, categoryName, seriesData);
		BarChartFillUtill.fillBarChartByChartName(slides.get(2), barChartData);
		
		//生成报告
		FileOutputStream outPutStream = new FileOutputStream(new File("E:\\POI测试用PPT\\BarChart测试生成.pptx"));
		ppt.write(outPutStream);
		outPutStream.close();
	}
}
