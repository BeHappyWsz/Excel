package com.wsz.Excel.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

public class PoiTest {

	public static void main(String[] args) {
		createExcel();
		
		readExcel();
	}
	
	/**
	 * 读取Excel文件内容
	 */
	@SuppressWarnings("resource")
	public static void readExcel() {
		//目标文件
		File file = new File("d:/test.xls");
		try {
			//创建Excel,读取文件内容
			HSSFWorkbook workBook = new HSSFWorkbook(FileUtils.openInputStream(file));
			//获取第一个工作表
//			HSSFSheet sheet = workBook.getSheet("Sheet0");
			HSSFSheet sheet = workBook.getSheetAt(0);
			int firstRowNum = 0;
			int lastRowNum = sheet.getLastRowNum();
			
			for(int i =firstRowNum;i<=lastRowNum;i++) {
				HSSFRow row = sheet.getRow(i);
				//获取当前行最后单元格列号
				short lastCellNum = row.getLastCellNum();
				for(int j =0;j<lastCellNum;j++) {
					HSSFCell cell = row.getCell(j);
					
					CellType type = cell.getCellTypeEnum();
					if(type == CellType.STRING) {
						String value = cell.getStringCellValue();
						System.out.print(value+" ");
					}else if(type == CellType.NUMERIC) {
						double value = cell.getNumericCellValue();
						System.out.print(value+" ");
					}
				}
				System.out.println();
			}
			
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}
	
	/**
	 * 创建Excel文件.xls(不能.xlsx)与内容
	 */
	@SuppressWarnings("resource")
	public static void createExcel() {
		String[] title = {"id","name","sex"};
		//创建Excel工作簿
		HSSFWorkbook workBook = new HSSFWorkbook();
		//创建一个工作表sheet
		HSSFSheet sheet = workBook.createSheet();
		//创建第一行
		HSSFRow row = sheet.createRow(0);
		HSSFCell cell = null;
		//插入第一行数据标题id,name,sex
		for(int i =0;i<title.length;i++) {
			cell = row.createCell(i);
			cell.setCellValue(title[i]);
		}
		//追加后续数据
		HSSFRow nextrow = null;
		HSSFCell cell2 = null;
		for(int i =1;i<=10;i++) {
			nextrow = sheet.createRow(i);
			
			cell2 = nextrow.createCell(0);
			cell2.setCellValue("a"+i);
			
			cell2 = nextrow.createCell(1);
			cell2.setCellValue("b"+i);
			
			cell2 = nextrow.createCell(2);
			cell2.setCellValue("男");
		}
		
		//创建一个文件
		File file = new File("d:/test.xls");
		try {
			file.createNewFile();
			FileOutputStream fos = FileUtils.openOutputStream(file);
			workBook.write(fos);
			fos.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
