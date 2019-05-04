package com.my.demo.excel;

import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author MaYue
 * @Version CreateTime:2019年4月5日下午10:11:59
 *用于替换数据
 */

public class ExcelFillUtill {
	
	public static void fillExcelByCellPosition(XSSFWorkbook excel,String sheetName
			, CellPositionData fillData) {
		
		XSSFSheet sheet = excel.getSheet(sheetName);
		
		Map<Integer, List<String>> data = fillData.getData();
		Set<Entry<Integer,List<String>>> entrySet = data.entrySet();
		for (Entry<Integer, List<String>> entry : entrySet) {
			
			XSSFRow row = sheet.getRow(entry.getKey());
			List<String> dataList = entry.getValue();
			if (dataList == null || dataList.isEmpty()) {
				continue;
			}
			//用于从list中取出数据
			int cellValueIndex = 0;
			
			for (int i = fillData.getStartColumn(); i < fillData.getEndColumn() + 1; i++) {
				//如果list中的数据少于单元格数量多于的就不填写
				if (cellValueIndex > dataList.size() - 1) {
					continue;
				}
				row.getCell(i).setCellValue(dataList.get(cellValueIndex));
				cellValueIndex++;
			}
			
		}
		
	}
}
