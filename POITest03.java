package com.atguigu.poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * 给单元格设置格式和列宽
 */
public class POITest03 {

	public static void main(String[] args) throws IOException {
		
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
		
		HSSFSheet hssfSheet = hssfWorkbook.createSheet("工作区01");
		
		HSSFRow hssfRow = hssfSheet.createRow(0);
		
		//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		HSSFDataFormat hssfDataFormat = hssfWorkbook.createDataFormat(); //创建Excel格式对象
		
		HSSFCellStyle dateCellStyle = hssfWorkbook.createCellStyle();
		dateCellStyle.setDataFormat(hssfDataFormat.getFormat("yyyy-MM-dd HH:mm:ss")); //设置日期格式
		
		HSSFCellStyle numberCellStyle = hssfWorkbook.createCellStyle();
		numberCellStyle.setDataFormat(hssfDataFormat.getFormat("#,#.00000")); //设置数值格式
		
		HSSFCellStyle wrapCellStyle = hssfWorkbook.createCellStyle();
		wrapCellStyle.setWrapText(true);//设置回绕文本风格样式，自动换行
		//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		
		//---------------------------------------
		HSSFCell cell00 = hssfRow.createCell(0);
		cell00.setCellValue("我的第一个单元格");
		
		HSSFCell cell01 = hssfRow.createCell(1);
		cell01.setCellValue(false);
		
		HSSFCell cell02 = hssfRow.createCell(2);
		cell02.setCellValue(Calendar.getInstance());
		cell02.setCellStyle(dateCellStyle);
		
		HSSFCell cell03 = hssfRow.createCell(3);
		cell03.setCellValue(new Date());
		cell03.setCellStyle(dateCellStyle);
		
		HSSFCell cell04 = hssfRow.createCell(4);
		cell04.setCellValue(123456789.87654321);
		cell04.setCellStyle(numberCellStyle);
		
		HSSFCell cell05 = hssfRow.createCell(5);
		cell05.setCellValue(new HSSFRichTextString("Rich Text Goooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooods"));
		cell05.setCellStyle(wrapCellStyle);
		//---------------------------------------
		
		//设置指定的列宽
		//hssfSheet.setColumnWidth(3, 8000);
		//hssfSheet.setColumnWidth(4, 8000);
		
		hssfSheet.setColumnWidth(5, 10000);
		
		//设置自动的列宽
		hssfSheet.autoSizeColumn(0);
		hssfSheet.autoSizeColumn(1);
		hssfSheet.autoSizeColumn(2);
		hssfSheet.autoSizeColumn(3);
		hssfSheet.autoSizeColumn(4);
		
		FileOutputStream fos = new FileOutputStream("good.xls");
		hssfWorkbook.write(fos);
	}

}
