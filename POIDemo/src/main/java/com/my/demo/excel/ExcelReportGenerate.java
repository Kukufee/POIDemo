package com.my.demo.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections4.map.HashedMap;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author MaYue
 * @Version CreateTime:2019年4月5日下午11:07:32
 * 具体的报告生成，一个报告对应一个类
 */

public class ExcelReportGenerate {
	
	public void generateTestExcel() throws IOException, InvalidFormatException {
		//1.拿到外面去
		XSSFWorkbook testWorkbook = new ExcelFactory().createTextExcel("target/excel/test.xlsx");
		//2.组装数据
		Map<Integer, List<String>> testData = new HashedMap<Integer,List<String>>();
		//第二行数据
		List<String> row2 = new ArrayList<String>();
		row2.add("昊然");
		row2.add("18");
		row2.add("2019-01-1-28");
		testData.put(1, row2);
		//第三行数据
		List<String> row3 = new ArrayList<String>();
		row3.add("吴磊");
		row3.add("18");
		row3.add("2019-01-1-29");
		testData.put(2, row3);
		
		CellPositionData cellPositionData = new CellPositionData(0, 2, testData);
		//替换数据
		ExcelFillUtill.fillExcelByCellPosition(testWorkbook, "Sheet1", cellPositionData);
		//生成报告
		FileOutputStream stream = new FileOutputStream(new File("target/report/test.xlsx"));
		testWorkbook.write(stream);
		stream.close();
	}
}
