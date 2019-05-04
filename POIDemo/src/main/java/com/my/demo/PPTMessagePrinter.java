package com.my.demo;

import java.util.List;

import org.apache.poi.sl.usermodel.TableCell;
import org.apache.poi.sl.usermodel.TableShape;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.sl.usermodel.TextRun;

public class PPTMessagePrinter {
	
	public static void printTextMessage(List<TextParagraph> textParagraphs) {
		for (TextParagraph textParagraph : textParagraphs) {
			List<TextRun> textRuns = textParagraph.getTextRuns();
			for (TextRun textRun : textRuns) {
				String rawText = textRun.getRawText();
				System.out.println("文本信息" + rawText);
			}
		}
	}
	
	public static void printTableMessage(TableShape tableShape) {
		
		int rowNumber = tableShape.getNumberOfRows();
		int columNumber = tableShape.getNumberOfColumns();
		
		for (int i = 0; i < rowNumber; i++) {
			for (int j = 0; j < columNumber; j++) {
				TableCell tableCell = tableShape.getCell(i, j);
				List<TextParagraph> textParagraphs = tableCell.getTextParagraphs();
				printMessage(textParagraphs);
			}
		}
	}
	
	private static void printMessage(List<TextParagraph> textParagraphs) {
		for (TextParagraph textParagraph : textParagraphs) {
			List<TextRun> textRuns = textParagraph.getTextRuns();
			for (TextRun textRun : textRuns) {
				String rawText = textRun.getRawText();
				System.out.println("表格信息" + rawText);
			}
		}
	}
}
