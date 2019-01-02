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
		//��ȡ�ù���������
		int sheets = workbook.getNumberOfSheets();
		System.out.println("����������:"+sheets);
		//����ȫ���Ĺ�����
		for (int i = 0; i < sheets; i++) {
			Sheet sheet = workbook.getSheetAt(i);
			//��ȡ����������
			String sheetName = sheet.getSheetName();
			System.out.println("����������:"+sheetName);
			//��ȡ�ù��������������ݵ���
			int rows = sheet.getPhysicalNumberOfRows();
			for (int j = 0; j < rows; j++) {
				Row row = sheet.getRow(j);
			//��ȡָ�������������ݵĵ�Ԫ������
				int cells = row.getPhysicalNumberOfCells();
				for (int k = 0; k < cells; k++) {
					Cell cell = row.getCell(k);
					//�жϵ�Ԫ������ͣ���Ӧ��ȡ����
					if (cell.getCellTypeEnum()==CellType.STRING) {
						//�����ַ�������
						System.out.print(cell.getStringCellValue()+"\t");
					} else if(cell.getCellTypeEnum()==CellType.NUMERIC){
						//������������ȡ
						//�������ָ�ʽ���Ĺ���
						DecimalFormat df = new DecimalFormat("####");
						System.out.print(df.format(cell.getNumericCellValue())+"\t");	
					}else if(cell.getCellTypeEnum()==CellType.BOOLEAN){
						//����boolean����
						System.out.print(cell.getBooleanCellValue()+"\t");
					}else if(cell.getCellTypeEnum()==CellType.BLANK){
						System.out.print("null\t");
					}else{
						//����ʱ���ʽ��
						System.out.print(cell.getDateCellValue()+"\t");
					}
					
				}
				System.out.println("");				
			}
			
		}
		workbook.close();
	}

}
