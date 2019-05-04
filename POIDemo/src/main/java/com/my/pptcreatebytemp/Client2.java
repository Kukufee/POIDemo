package com.my.pptcreatebytemp;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

/**
 *
 * @author MaYue
 * @Version CreateTime:2019年4月6日下午11:38:20
 *
 */

public class Client2 {
	
	public static void main(String[] args) throws IOException {
		XMLSlideShow ppt = PPTCreateUtill.createPPT("E:\\POI测试用PPT\\POI测试.pptx");
		//获取所有页数，页数顺序为list元素顺序
		List<XSLFSlide> slides = PPTCreateUtill.getSlides(ppt);
		//按页填充
		//第一页
		ArrayList<String> pageOneTextData = new ArrayList<String>();
		pageOneTextData.add("第一页文本替换01");
		ArrayList<String> pageOneTableData = new ArrayList<String>();
		pageOneTableData.add("第一页表格替换01");
		pageOneTableData.add("第一页表格替换02");
		pageOneTableData.add("第一页表格替换03");
		
		PPTCreateUtill.fillTextData(slides.get(0), pageOneTextData);
		PPTCreateUtill.fillTableData(slides.get(0), pageOneTableData);

		//第二页
		ArrayList<String> pageTwoTextData = new ArrayList<String>();
		pageTwoTextData.add("第二页文本替换01");
		pageTwoTextData.add("第二页文本替换02");
		ArrayList<String> pageTwoTableData = new ArrayList<String>();
		pageTwoTableData.add("第二页表格替换01");
		pageTwoTableData.add("第二页表格替换02");
		pageTwoTableData.add("第二页表格替换03");
		pageTwoTableData.add("第二页表格替换04");
		
		PPTCreateUtill.fillTextData(slides.get(1), pageTwoTextData);
		PPTCreateUtill.fillTableData(slides.get(1), pageTwoTableData);
		
		//生成报告
		FileOutputStream outPutStream = new FileOutputStream(new File("E:\\POI测试用PPT\\POI测试生成.pptx"));
		ppt.write(outPutStream);
		outPutStream.close();
	}
}
