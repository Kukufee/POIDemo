package com.my.ppt.chart.client;

import java.io.IOException;
import java.util.List;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import com.my.ppt.chart.ChartIndexPrintUtil;
import com.my.pptcreatebytemp.PPTCreateUtill;

/**
 *
 * @author MaYue
 * @Version CreateTime:2019年4月8日下午10:26:02
 * 使用ChartIndexGetUtil获取Chart对应的下标
 */

public class ClientGetTempChartIndexAndTitle {
	public static void main(String[] args) throws IOException {
		//路径可改
		XMLSlideShow ppt = PPTCreateUtill.createPPT("E:\\POI测试用PPT\\POI测试.pptx");
		//获取所有页数，页数顺序为list元素顺序
		List<XSLFSlide> slides = PPTCreateUtill.getSlides(ppt);
		ChartIndexPrintUtil.printIndexAndTitle(slides.get(2));
	}
}
