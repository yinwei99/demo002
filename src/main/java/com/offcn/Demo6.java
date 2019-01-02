package com.offcn;

import java.io.File;
import java.text.DecimalFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Demo6 {

	public static void main(String[] args) throws Exception {
		Workbook workbook = WorkbookFactory.create(new File("d:\\chart\\text.xlsx"));
		//获取该工作表数量
		int sheets = workbook.getNumberOfSheets();
		System.out.println("工作表数量:"+sheets);
		//遍历全部的工作表
		for (int i = 0; i < sheets; i++) {
			Sheet sheet = workbook.getSheetAt(i);
			//获取工作表名称
			String sheetName = sheet.getSheetName();
			System.out.println("工作表名称:"+sheetName);
			//获取该工作包包含有数据的行
			int rows = sheet.getPhysicalNumberOfRows();
			for (int j = 0; j < rows; j++) {
				Row row = sheet.getRow(j);
			//获取指定行里面有数据的单元格数量
				int cells = row.getPhysicalNumberOfCells();
				for (int k = 0; k < cells; k++) {
					Cell cell = row.getCell(k);
					//判断单元格的类型，对应读取数据
					if (cell.getCellTypeEnum()==CellType.STRING) {
						//按照字符串来读
						System.out.print(cell.getStringCellValue()+"\t");
					} else if(cell.getCellTypeEnum()==CellType.NUMERIC){
						//按照数字来读取
						//创建数字格式化的工具
						DecimalFormat df = new DecimalFormat("####");
						System.out.print(df.format(cell.getNumericCellValue())+"\t");	
					}else if(cell.getCellTypeEnum()==CellType.BOOLEAN){
						//按照boolean来读
						System.out.print(cell.getBooleanCellValue()+"\t");
					}else if(cell.getCellTypeEnum()==CellType.BLANK){
						System.out.print("null\t");
					}else{
						//按照时间格式读
						System.out.print(cell.getDateCellValue()+"\t");
					}
					
				}
				System.out.println("");				
			}
			
		}
		workbook.close();
	}

}
