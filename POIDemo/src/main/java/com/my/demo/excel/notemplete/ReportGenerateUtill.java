package com.my.demo.excel.notemplete;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author MaYue
 * @Version CreateTime:2019年4月6日上午10:56:00
 *
 */

public class ReportGenerateUtill {
	
	public static void createExcelReport(XSSFWorkbook workBook, String repotPath) throws IOException {
		FileOutputStream stream = new FileOutputStream(new File(repotPath));
		workBook.write(stream);
		stream.close();
	}
	
	/**
	 * 
	 * @param data 需要填充的数据，map代表行数据
	 * @param header 表头信息
	 * @param workbook
	 * @param sheetName
	 * 填充单个sheet
	 * 表数据都为String，需要Sheet名和表头名
	 * 多线程分别生成不同的workbook即可
	 */
	public static void fillExcel(XSSFWorkbook workBook, String sheetName
			, String[] header, List<Map<String,String>> data) {
		
		//填写表头数据
		XSSFSheet sheet = workBook.createSheet(sheetName);
		XSSFRow row0 = sheet.createRow(0);
		for (int i = 0; i < header.length; i++) {
			XSSFCell cell = row0.createCell(i);
			//一般使用date，double，String三种数据类型
			cell.setCellValue(header[i]);
		}
		
		//1.整理数据，因为map是无序的
		//整个sheet的数据
		List<List<String>> sheetData = new ArrayList<List<String>>();
		for (int i = 0; i < data.size(); i++) {
			//根据map将行数据存入list集合
			//行数据
			List<String> rowlist = new ArrayList<String>();
			Map<String, String> rowData = data.get(i);
			for (int j = 0; j < header.length; j++) {
				rowlist.add(rowData.get(header[j]));
			}
			
			sheetData.add(rowlist);
		}
		
		//2.使用sheetData填充创建行数据，从第二行开始填
		for (int i = 0; i < sheetData.size(); i++) {
			XSSFRow row = sheet.createRow(i + 1);
			//填充cell
			List<String> rowData = sheetData.get(i);
			for (int j = 0; j < rowData.size(); j++) {
				//全部转化成string类型
				row.createCell(j).setCellValue(rowData.get(j));
			}
		}
		
	}
}
