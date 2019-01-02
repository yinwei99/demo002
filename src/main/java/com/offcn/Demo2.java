package com.offcn;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Demo2 {
	public static void main(String[] args) throws IOException {
		
	FileInputStream stream = new FileInputStream("D:\\chart\\demo1.xls");
	
	HSSFWorkbook workbook = new HSSFWorkbook(stream);
	HSSFSheet sheet = workbook.getSheet("工作表1");
	HSSFRow row = sheet.getRow(0);
	HSSFCell cell = row.getCell(3);
	String value = cell.getStringCellValue();
	System.out.println("读取："+value);
	
		
	}
}
