package com.mycompany.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import Utility.ExcelReadUtil;

public class ExcelRead {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		String path = "C:\\Users\\Syan\\Desktop\\MySheet.xlsx";
		ExcelReadUtil exceldata = new ExcelReadUtil(path);
		exceldata.writeData("Sheet1");
		
		
			}

}
