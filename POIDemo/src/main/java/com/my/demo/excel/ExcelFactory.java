package com.my.demo.excel;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author MaYue
 * @Version CreateTime:2019年4月5日下午9:54:41
 *此类用于读取excel文件，创建workbook
 */

public class ExcelFactory {
	
	public XSSFWorkbook createTextExcel(String filePath) throws IOException, InvalidFormatException {
		return getTextExcel(filePath);
	}
	
	private XSSFWorkbook getTextExcel(String filePath) throws IOException, InvalidFormatException {
		File file = new File(filePath);
		XSSFWorkbook testExcel = new XSSFWorkbook(file);
		return testExcel;
	}
	
}
