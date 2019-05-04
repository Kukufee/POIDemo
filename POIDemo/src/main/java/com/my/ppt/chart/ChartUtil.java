package com.my.ppt.chart;

import java.util.List;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;

/**
 *
 * @author MaYue
 * @Version CreateTime:2019年4月7日下午5:38:37
 * chart操作的工具类
 */
public class ChartUtil {
	
	/**
	 * 
	 * @param slide
	 * @return
	 * 图表的title必须放在第一个位置，且图表标题不能重复
	 */
	public static XSLFChart getChartByTitle(XSLFSlide slide, String chartTitle) {
		XSLFChart chart = null;
		List<POIXMLDocumentPart> parts = slide.getRelations();
		if (parts == null || parts.size() == 0) {
			return chart;
		}
		for (POIXMLDocumentPart part : parts) {
			if (part instanceof XSLFChart) {
				XSLFChart chartPart = (XSLFChart) part;
				//获取chart的文本第一段，标题在第一段中
				String templetTitle = "";
				CTTextBody body = chartPart.getCTChart().getTitle().getTx().getRich();
				//String templetTitle = body.getPArray(0).getRArray(0).getT();
				List<CTTextParagraph> pList = body.getPList();
				for (CTTextParagraph titleParagraph : pList) {
					List<CTRegularTextRun> rList = titleParagraph.getRList();
					for (CTRegularTextRun titleLine : rList) {
						if (!titleLine.getT().isEmpty()) {
							templetTitle += titleLine.getT();
						}
					}
				}
				System.out.println(templetTitle);
				if (chartTitle.equals(templetTitle.trim())) {
					chart = chartPart;
					return chart;
				}
			}
		}
		return chart;
	}
	
}
