package com.poi.test;

import java.io.FileOutputStream;

import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Before;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ExcelTest {
	private static Logger log = LoggerFactory.getLogger(ExcelTest.class);
	
	@Before
	public void before(){
		log.debug("before");
	}
	
	@Test
	public void testPoiExcel() throws Throwable {
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet();
		for(int i=0; i<10; i++){
			Row row = sheet.createRow(i);
			for(int j=0; j<5; j++){
				Cell cell = row.createCell(j);
				cell.setCellValue(j+"|你好|"+i);
			}
		}
		FileOutputStream fileOut = new FileOutputStream("workbook.xls");
		wb.write(fileOut);
		fileOut.close();
	}
	
	@Test
	public void testJxlExcel() throws Exception{
		FileOutputStream fileOut = new FileOutputStream("workbook2.xls");
		WritableWorkbook wwb = jxl.Workbook.createWorkbook(fileOut);
		// 生成工作表,(name:First Sheet,参数0表示这是第一页)
		WritableSheet sheet = wwb.createSheet("First Sheet", 0);
		 // 开始写入第一行(即标题栏)
		String[] title = {"姓名","性别","年龄"};
        for (int i=0; i<title.length; i++) {
            // 用于写入文本内容到工作表中去
            // 在Label对象的构造中指明单元格位置(参数依次代表列数、行数、内容 )      
            Label label = new Label(i, 0, title[i]);
            // 将定义好的单元格添加到工作表中
            sheet.addCell(label);
        }
        
        // 开始写入内容
        for (int row=0; row<10; row++) {        
             // 数据是文本时(用label写入到工作表中)
             for (int col=0; col<3; col++) {                
            	 String value = "小李"+col;
            	 if(col == 1)
            		 value = "男";
            	 if(col == 2)
            		 value = "2"+row;
            	 Label label = new Label(col, row+1, value);
                 sheet.addCell(label);            
             }   
        }
        
        // 写入数据
        wwb.write();
        // 关闭文件
        wwb.close();
        // 关闭输出流
        fileOut.close();
	}
}
