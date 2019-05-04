package com.my.demo.excel;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 *
 * @author MaYue
 * @Version CreateTime:2019年4月5日下午11:32:35
 *
 */

public class Client {
	public static void main(String[] args) throws IOException, InvalidFormatException {
		ExcelReportGenerate generate = new ExcelReportGenerate();
		generate.generateTestExcel();
		System.out.println("已经生成报告");
	}
}
