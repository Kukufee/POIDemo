package com.my.demo;

import java.io.File;
import java.util.List;

import org.apache.poi.sl.usermodel.TableShape;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.sl.usermodel.TextShape;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class PPTFiller {
	public static void main(String[] args) {
		
		File pptFile = new File("E:\\POI测试用PPT\\POI测试.pptx");
		PPTReader pptReader = new PPTReader(pptFile);
		XMLSlideShow ppt = pptReader.getSlideShow();
		
		//获取ppt每页对象的集合
		List<XSLFSlide> slides = ppt.getSlides();
		System.out.println("ppt页数:" + slides.size());
		
		//获取每页ppt
		for (int i = 0; i < slides.size(); i++) {
			int pptPage = i + 1;
			System.out.println("第" + pptPage + "页");
			
			XSLFSlide slide = slides.get(i);
			//获取每页的每个模块：文本、表格、图表等
			List<XSLFShape> shapes = slide.getShapes();
			for (int j = 0; j < shapes.size(); j++) {
				XSLFShape shape = shapes.get(j);
				//对shap进行处理
				//处理文本
				if (shape instanceof TextShape) {
					List<TextParagraph> textParagraphs = ((TextShape) shape).getTextParagraphs();
					//输出文本内容
					PPTMessagePrinter.printTextMessage(textParagraphs);
				}
				
				//处理表格
				if (shape instanceof TableShape) {
					PPTMessagePrinter.printTableMessage((TableShape)shape);
				}
				
			}
			
		}
		
	}
}
