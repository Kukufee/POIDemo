package com.my.pptcreatebytemp;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.sl.usermodel.TableCell;
import org.apache.poi.sl.usermodel.TableShape;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.sl.usermodel.TextRun;
import org.apache.poi.sl.usermodel.TextShape;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

/**
 *
 * @author MaYue
 * @Version CreateTime:2019年4月6日下午6:44:15
 * 根据模板创建ppt、获取每页ppt、填充Text内容、填充table中Text内容
 */

public class PPTCreateUtill {
	
	/**
	 * @param relativPath
	 * @return
	 * @throws IOException
	 * 获取XMLSlideShow即ppt对象
	 */
	public static XMLSlideShow createPPT(String path) throws IOException {
		FileInputStream fileInputStream = new FileInputStream(new File(path));
		XMLSlideShow ppt = new XMLSlideShow(fileInputStream);
		fileInputStream.close();
		return ppt;
	}
	
	/**
	 * @param ppt
	 * @return 通过ppt获取slids,页序号与集合顺序一致
	 * @throws IOException 
	 */
	public static List<XSLFSlide> getSlides(XMLSlideShow ppt) throws IOException {
		List<XSLFSlide> slides = ppt.getSlides();
		return slides;
	}
	
	/**
	 * 
	 * @param slide
	 * @param replaceData 用于替换标识为${0}...${n}
	 * 顺序与集合一致，数量也必须一致
	 */
	public static void fillTextData(XSLFSlide slide, List<String> replaceData) {
		
		if (replaceData == null || replaceData.isEmpty()) {
			return;
		}
		
		List<XSLFShape> shapes = slide.getShapes();
		for (XSLFShape shape : shapes) {
			//如果为文本则执行文本替换操作
			if (shape instanceof TextShape) {
				//替换标识${0}...${n}
				String shapeText = ((TextShape) shape).getText();
				//如果为空则跳过当前循环
				if (shapeText.isEmpty()) {
					continue;
				}
				
				List<TextParagraph> paras = ((TextShape) shape).getTextParagraphs();
				//替换标记内容
				replaceTextRun(paras, replaceData);
				}
		}
	}
	
	/**
	 * 
	 * @param slide
	 * @param replaceData 用于替换标识为${0}...${n}，与table的标识分开，
	 * 顺序与集合一致，数量也必须一致
	 */
	public static void fillTableData(XSLFSlide slide, List<String> replaceData) {
		List<XSLFShape> shapes = slide.getShapes();
		
		if (shapes == null || shapes.isEmpty()) {
			return;
		}
		
		for (XSLFShape shape : shapes) {
			if (shape instanceof TableShape) {
				
				int rowNum = ((TableShape) shape).getNumberOfRows();
				int columnNum = ((TableShape) shape).getNumberOfColumns();
				for (int i = 0; i < rowNum; i++) {
					for (int j = 0; j < columnNum; j++) {
						
						TableCell cell = ((TableShape) shape).getCell(i, j);
						//替换标识${0}...${n}
						String cellText = cell.getText();
						//如果为空则跳过当前循环
						if (cellText.isEmpty()) {
							continue;
						}
						
						List<TextParagraph> paras = cell.getTextParagraphs();
						replaceTextRun(paras, replaceData);
						/*
						StringBuffer text = new StringBuffer(cellText);
						for (int z = 0; z < replaceData.size(); z++) {
							String sign = "$" + "{" + z + "}";
							//判断是否包含该标记
							//此方法：如果有则返回第一个字符的下标，如果没有则返回-1
							int index = text.indexOf(sign);
							if (index < 0) {
								continue;
							}
							text.replace(index, index + 4, replaceData.get(z));
						}
						cell.setText(text.toString());*/
						System.out.println(cell.getText());
					}
				}
			}
		}
		
		
	}
	
	/**
	 * 
	 * @param paras
	 * @param replaceData
	 * 替换标记内容
	 */
	private static void replaceTextRun(List<TextParagraph> paras, List<String> replaceData) {
		for (TextParagraph paragraph : paras) {
			
			List<TextRun> textRuns = paragraph.getTextRuns();
			for (TextRun textRun : textRuns) {
				StringBuffer text = new StringBuffer(textRun.getRawText());
				for (int i = 0; i < replaceData.size(); i++) {
					String sign = "$" + "{" + i + "}";
					//判断是否包含该标记
					//此方法：如果有则返回第一个字符的下标，如果没有则返回-1
					int index = text.indexOf(sign);
					if (index < 0) {
						continue;
					}
					text.replace(index, index + 4, replaceData.get(i));
					textRun.setText(text.toString());	
				}
			}
		}	
	}
}

