package com.offcn;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo3 {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("������1");
		XSSFRow row = sheet.createRow(0);
		XSSFCell cell = row.createCell(3);
		cell.setCellValue("�°湤����");
		workbook.write(new FileOutputStream("d:\\chart\\demo3.xlsx"));
		System.out.println("�����ɳɹ�!");
		workbook.close();

	}

}
