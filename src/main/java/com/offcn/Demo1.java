package com.offcn;

import java.io.File;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Demo1 {

	public static void main(String[] args) throws IOException {
		// 1.��������������
		HSSFWorkbook workbook = new HSSFWorkbook();
		//
		HSSFSheet sheet = workbook.createSheet("������1");
		HSSFRow row = sheet.createRow(0);
		HSSFCell cell = row.createCell(3);
		cell.setCellValue("��ã�");
		workbook.write(new File("d:\\chart\\demo1.xls"));
		workbook.close();
		
	}

}
