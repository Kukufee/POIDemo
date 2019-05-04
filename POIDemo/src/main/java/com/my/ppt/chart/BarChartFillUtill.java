package com.my.ppt.chart;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumVal;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrVal;

/**
 *
 * @author MaYue
 * @Version CreateTime:2019年4月7日下午3:06:27
 * 
 */
public class BarChartFillUtill {
	
	/**
	 * 
	 * @param slide
	 * @param chartData
	 * 1.根据Title生成bar
	 * 2.填充bar的excel（workbook）并设置
	 * 3.指定barchart的数据显示
	 */
	public static void fillBarChartByChartName(XSLFSlide slide, BarChartData chartData) {
		
		//填充图表，根据Title获取chart
		XSLFChart chart = ChartUtil.getChartByTitle(slide, chartData.title);
		if (chart == null) {
			return;
		}
		fillBarChart(chart, chartData);
	}
	
	public static void fillBarChart(XSLFChart chart, BarChartData chartData) {
		//获取数据
		String[] seriesName = chartData.series;
		String[] categorys = chartData.categoryName;
		double[][] serieData = chartData.serieData;
		int catCount = categorys.length;
		
		//创建一个柱状图对应的内置excel
		XSSFWorkbook chartWorkBook = new XSSFWorkbook();
		XSSFSheet chartSheet = fillChartWorkBook(chartWorkBook, chartData);
		//填写chart,按照每个serie填写
		CTPlotArea plotArea = chart.getCTChart().getPlotArea();
		CTBarChart ctBarChart = plotArea.getBarChartList().get(0);
		List<CTBarSer> serList = ctBarChart.getSerList();
		int tempSeriesNum = serList.size();
		int refreshSeriesNum = seriesName.length;
		//判断模板系列是否匹配
		//需要系列数量小于模板系列数量
		if (refreshSeriesNum < tempSeriesNum) {
			//需要减去的系列数量
			int niuesNum = tempSeriesNum - refreshSeriesNum;
			for (int i = 0; i < niuesNum; i++) {
				serList.remove(i);
			}
		}
		
		for (int i = 0; i < refreshSeriesNum; i++) {
			//系列名与sheet设置链接进行对应
			CTBarSer ser = serList.get(i);
			CTSerTx serTx = ser.getTx();
			//填充该serie的名字
			serTx.getStrRef().getStrCache().getPtList()
				.get(0).setV(seriesName[i]);
			//设置图表Title，每个Ser都设置一遍吗？
			/*sheet.getRow(0).getCell(i + 1)
		        .setCellValue(seriesDataList.get(0).getSeriesName());*/
			//系列名与sheet设置链接进行对应
			String serString = new CellReference(chartSheet.getSheetName(), 0, i + 1, true, true).formatAsString();
			serTx.getStrRef().setF(serString);
			//填充该系列下的图例数据（category），分为图例名称和图例数据
			//series数据下标从0开始
			CTStrRef strRef = ser.getCat().getStrRef();
			CTStrData catName = strRef.getStrCache();
			CTNumRef numRef = ser.getVal().getNumRef();
			CTNumData catNumData = numRef.getNumCache();
			
			catName.setPtArray(null);  // unset old axis text 
			catNumData.setPtArray(null);  // unset old values 
			//填充catName
			for (int j = 0; j < catCount; j++) {
				CTStrVal strVal = catName.addNewPt();
				strVal.setIdx(j);
				strVal.setV(categorys[j]);
			}
			//填充catNumData
			for (int k = 0; k < catCount; k++) {
				CTNumVal numVal = catNumData.addNewPt();
				numVal.setIdx(k);
				numVal.setV(String.valueOf(serieData[i][k]));
			}
			
			//count以idx的最后一位为准（从0开始），即category数量-1
			catName.getPtCount().setVal(catCount - 1);
			catNumData.getPtCount().setVal(catCount - 1);
			
			//设置sheet和图表创立连接的系列的数据位置
			//设置图裂num数据的范围，第二行第二列开始到最后一行第二列
			String numRangeAddress = new CellRangeAddress(1, chartSheet.getLastRowNum(), i + 1, i + 1)
					.formatAsString(chartSheet.getSheetName(), true);
			numRef.setF(numRangeAddress);
			//创建横轴连接,categoryName
			String strRangeAddress = new CellRangeAddress(1, chartSheet.getLastRowNum(), 0, 0)
					.formatAsString(chartSheet.getSheetName(), true);
			strRef.setF(strRangeAddress);
		}
		//chart.getR
		POIXMLDocumentPart docPart = chart.getRelations().get(0);
		OutputStream outStream = docPart.getPackagePart().getOutputStream();
		try {
			chartWorkBook.write(outStream);
			outStream.close();
		} catch (IOException e) {
			e.printStackTrace();
				}
	}
	/**
	 * 
	 * @param part
	 * @return
	 * 判断图表是什么类型的
	 */
	private String chartTypeCheck(XSLFChart part) {
		String chartType = "";
        CTPlotArea plotArea = part.getCTChart().getPlotArea();
        if (plotArea.getLineChartList().size() != 0) {
            chartType = "lie";
        }
        
        if (plotArea.getBarChartList().size() != 0) {
            chartType = "bar";
        }
        if (plotArea.getLineChartList().size() != 0
                && plotArea.getBarChartList().size() != 0) {
            chartType = "barAndlie";
        }
        if (plotArea.getPieChartList().size() != 0) {
            chartType = "pie";
        }

            return chartType;
	}
	
	private static XSSFSheet fillChartWorkBook(XSSFWorkbook chartWorkBook, BarChartData chartData) {
		
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
	
	
	
	
}
