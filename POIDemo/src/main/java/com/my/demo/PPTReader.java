package com.my.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

public class PPTReader {
	
	private File pptFile;
	
	public PPTReader(File pptFile) {
		this.pptFile = pptFile;
	}

	public PPTReader() {
		
	}

	public XMLSlideShow getSlideShow() {
		FileInputStream pptInputStream = null;
		try {
			pptInputStream = new FileInputStream(pptFile);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return null;
		}
		
		try {
			XMLSlideShow ppt = new XMLSlideShow(pptInputStream);
			return ppt;
		} catch (IOException e) {
			e.printStackTrace();
			return null;
		}
	}
	
}
