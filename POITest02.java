package com.atguigu.poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * 单元格设置不同类型的值
 * @author zhangyu
 *
 */
public class POITest02 {

	public static void main(String[] args) throws IOException {
		
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
		
		HSSFSheet hssfSheet = hssfWorkbook.createSheet("工作区01");
		
		HSSFRow hssfRow = hssfSheet.createRow(0);
		
		//---------------------------------------
		HSSFCell cell00 = hssfRow.createCell(0);
		cell00.setCellValue("我的第一个单元格");
		
		HSSFCell cell01 = hssfRow.createCell(1);
		cell01.setCellValue(false);
		
		HSSFCell cell02 = hssfRow.createCell(2);
		cell02.setCellValue(Calendar.getInstance());
		
		HSSFCell cell03 = hssfRow.createCell(3);
		cell03.setCellValue(new Date());
		
		HSSFCell cell04 = hssfRow.createCell(4);
		cell04.setCellValue(123456789.87654321);
		
		HSSFCell cell05 = hssfRow.createCell(5);
		cell05.setCellValue(new HSSFRichTextString("Rich Text"));
		//---------------------------------------
		
		FileOutputStream fos = new FileOutputStream("good.xls");
		hssfWorkbook.write(fos);
	}

}
