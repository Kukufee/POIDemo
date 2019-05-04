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
 * @Version CreateTime:2019年4月8日下午10:17:28
 * 通过该工具输出图表与之对应的index
 * 使用前提，图表必须有Title，获得下标后删除Title即可
 */
public class ChartIndexPrintUtil {
	public static void printIndexAndTitle(XSLFSlide slide) {
		//该下标为Chart在Relation中的位置
		List<POIXMLDocumentPart> parts = slide.getRelations();
		if (parts == null || parts.size() == 0) {
			System.out.println("没有图表");
		}
		for (int i = 0; i < parts.size(); i++) {
			if (parts.get(i) instanceof XSLFChart) {
				XSLFChart chart = (XSLFChart) parts.get(i);
				//获取路线Title-body-paragraph集合-line集合-text
				String templetTitle = "";
				CTTextBody body = chart.getCTChart().getTitle().getTx().getRich();
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
				System.out.println("chartIndex=" + i + "," + "chartname=" + templetTitle);
			}
		}
	}
}
