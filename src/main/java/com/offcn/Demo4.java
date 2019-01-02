package com.offcn;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo4 {

	public static void main(String[] args) throws IOException {
		FileInputStream stream = new FileInputStream("d:\\chart\\demo3.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook(stream);
		XSSFSheet sheet = workbook.getSheet("工作表1");
		XSSFRow row = sheet.getRow(0);
		XSSFCell cell = row.getCell(3);
		String value = cell.getStringCellValue();
		System.out.println("读取："+value);
		
		
		
		
		
	}
}
