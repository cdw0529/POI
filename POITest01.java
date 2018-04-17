package com.fhqb.poi;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * 第一个简单的POI编程
 * @author zhangyu
 */
public class POITest01 {

	public static void main(String[] args) throws IOException {
		
		//①创建Excel文件对象:HSSFWorkbook
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
		
		//②创建工作区
		HSSFSheet hssfSheet = hssfWorkbook.createSheet("工作区01");
		
		//③创建行
		HSSFRow hssfRow = hssfSheet.createRow(0);
		
		//④创建单元格
		HSSFCell cell = hssfRow.createCell(5);
		cell.setCellValue("我的第一个单元格");
		
		//⑤生成Excel文件
		FileOutputStream fos = new FileOutputStream("good.xls");
		hssfWorkbook.write(fos);
	}

}
