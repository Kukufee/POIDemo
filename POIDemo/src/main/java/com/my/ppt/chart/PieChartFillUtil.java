package com.my.ppt.chart;

import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumVal;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrVal;

/**
 *
 * @author MaYue
 * @Version CreateTime:2019年4月8日下午6:35:17
 *
 */

public class PieChartFillUtil {
	
	public static void fillPieChart(XSLFSlide slide, PieChartData chartData) {
		//填充图表，根据Title获取chart
		XSLFChart chart = ChartUtil.getChartByTitle(slide, chartData.title);
		if (chart == null) {
			return;
		}
		fillPieChartByData(chart, chartData);
	}
	
	public static void fillPieChartByData(XSLFChart chart, PieChartData chartData) {
		//创建一个柱状图对应的内置excel
		XSSFWorkbook chartWorkBook = new XSSFWorkbook();
		XSSFSheet chartSheet = fillChartWorkBook(chartWorkBook, chartData);
		//填写chart,按照每个serie填写
		CTPlotArea plotArea = chart.getCTChart().getPlotArea();
		CTPieChart pieChart = plotArea.getPieChartList().get(0);
		CTPieSer pieSer = pieChart.getSerList().get(0);
		
		//1.操作series,填充series名，serie分为cat（strData图例名）和value（numData图例值）
		CTSerTx serTx = pieSer.getTx();
		//系列名称数组
		String seriesName = chartData.series[0];
		fillPieChartSerName(serTx, seriesName);
		serNamesLinkSheet(serTx, chartSheet);
		
		//2.填充series的cat
		//图例名称数组
		String[] categorys = chartData.categoryName;
		CTAxDataSource serCat = pieSer.getCat();
		fillPieChartSerStrData(serCat, categorys);
		//为图例名和sheet创立连接
		serStrDataLinkSheet(serCat, chartSheet);
		
		//3.填充series的val
		//图例值数组，对应excel的一列，竖向填写
		double[] serieData = chartData.serieData[0];
		CTNumDataSource serVal = pieSer.getVal();
		fillPieChartSerNumData(serVal, serieData);
		//为图例名和sheet创立连接
		serNumDataLinkSheet(serVal, chartSheet);
		
		//4.输出
		POIXMLDocumentPart docPart = chart.getRelations().get(0);
		OutputStream outStream = docPart.getPackagePart().getOutputStream();
		try {
			chartWorkBook.write(outStream);
			outStream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	private static XSSFSheet fillChartWorkBook(XSSFWorkbook chartWorkBook, PieChartData chartData) {
		
		//创建sheet
		XSSFSheet chartSheet = chartWorkBook.createSheet();
		//填写sheet的数据,根据行填写
		int rowNum = chartData.categoryName.length + 1;
		int colNum = chartData.series.length + 1;
		//先填写第一行，第一个cell不填内容
		XSSFRow seriesRow = chartSheet.createRow(0);
		for (int i = 1; i < colNum; i++) {
			seriesRow.createCell(i).setCellValue((chartData.series)[i-1]);
		}
		//填写第一列,第一个cell不填内容
		for (int i = 1; i < rowNum; i++) {
			XSSFRow row = chartSheet.createRow(i);
			row.createCell(0).setCellValue((chartData.categoryName)[i-1]);
		}
		
		//填写categorydata,一列一列的填
		for (int i = 1; i < colNum; i++) {
			for (int j = 1; j < rowNum; j++) {
				//填写excel时需要使用double数据，不然ppt点开excel是文本类型，数据会销失
				chartSheet.getRow(j).createCell(i).setCellValue(chartData.serieData[i - 1][j - 1]);
			}
		}
		return 	chartSheet;	
	}
	
	private static void fillPieChartSerName(CTSerTx serTx, String seriesName) {
		//填充该serie的名字
		serTx.getStrRef().getStrCache().getPtList()
			.get(0).setV(seriesName);
	}
	
	private static void serNamesLinkSheet(CTSerTx serTx, XSSFSheet chartSheet) {
		//系列名与sheet设置链接进行对应
		String serString = new CellReference(chartSheet.getSheetName(), 0, 1, true, true).formatAsString();
		serTx.getStrRef().setF(serString);
	}
	
	private static void fillPieChartSerStrData(CTAxDataSource serCat, String[] categorys) {
		//series数据下标从0开始
		CTStrData catNames = serCat.getStrRef().getStrCache();
		catNames.setPtArray(null);  // unset old axis text 
		//填充catName
		int catCount = categorys.length;
		for (int j = 0; j < catCount; j++) {
			CTStrVal strVal = catNames.addNewPt();
			strVal.setIdx(j);
			strVal.setV(categorys[j]);
		}
		//count以idx的最后一位为准（从0开始），即category数量-1
		catNames.getPtCount().setVal(catCount - 1);
	}
	
	private static void serStrDataLinkSheet(CTAxDataSource serCat, XSSFSheet chartSheet) {
		//创建横轴连接,categoryName
		String strRangeAddress = new CellRangeAddress(1, chartSheet.getLastRowNum(), 0, 0)
				.formatAsString(chartSheet.getSheetName(), true);
		serCat.getStrRef().setF(strRangeAddress);
	}
	
	private static void fillPieChartSerNumData(CTNumDataSource serVal, double[] serieDatas) {
		
		CTNumData numData = serVal.getNumRef().getNumCache();
		numData.setPtArray(null);  // unset old values 
		//填充numData
		int dataCount = serieDatas.length;
		for (int i = 0; i < dataCount; i++) {
			CTNumVal numVal = numData.addNewPt();
			numVal.setIdx(i);
			numVal.setV(String.valueOf(serieDatas[i]));
		}
		numData.getPtCount().setVal(dataCount - 1);
	}
	
	private static void serNumDataLinkSheet(CTNumDataSource serVal, XSSFSheet chartSheet) {
		//为serie的值即numdata与sheet创建链接
		//设置图裂num数据的范围，第二行第二列开始到最后一行第二列
		String numRangeAddress = new CellRangeAddress(1, chartSheet.getLastRowNum(), 1, 1)
				.formatAsString(chartSheet.getSheetName(), true);
		serVal.getNumRef().setF(numRangeAddress);
	}
}
