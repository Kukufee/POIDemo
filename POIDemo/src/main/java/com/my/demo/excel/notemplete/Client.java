package com.my.demo.excel.notemplete;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Map;

import org.apache.commons.collections4.map.HashedMap;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author MaYue
 * @Version CreateTime:2019年4月6日上午11:06:09
 *
 */

public class Client {
	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workBook = new XSSFWorkbook();
		String sheetName = TestReportSheetsAndHeader.SHEETS_NAME[0];
		String[] header = TestReportSheetsAndHeader.SHEET1_HEADER;
		
		ArrayList<Map<String, String>> sheetData = new ArrayList<Map<String,String>>();
		for (int i = 0; i < 5001; i++) {
			Map<String, String> rowData = new HashedMap<String,String>();
			rowData.put("name", "昊然");
			rowData.put("age", "18");
			rowData.put("gender", "男");
			sheetData.add(rowData);
		}
		ReportGenerateUtill.fillExcel(workBook, sheetName, header, sheetData);
		ReportGenerateUtill.createExcelReport(workBook, "target/report/test2.xlsx");
		System.out.println("生成报告成功");
	}
}
