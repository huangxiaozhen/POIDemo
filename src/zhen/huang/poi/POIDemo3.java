package zhen.huang.poi;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class POIDemo3
{
	
	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException
	{
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet("第一页sheet");
		Row row = sheet.createRow(0);
		Cell cell1 = row.createCell(0);
		cell1.setCellValue(1);
		
		row.createCell(1).setCellValue(false);
		row.createCell(2).setCellValue(1.2);
		row.createCell(3).setCellValue("这是一个字符串");
		
		FileOutputStream fileOutputStream = new FileOutputStream("D:\\单元格.xls");
		
		workbook.write(fileOutputStream);
		
		fileOutputStream.close();
	}

}
