package zhen.huang.poi;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

//单元格对齐样式

public class POIDemo7
{
	public static void main(String[] args) throws IOException
	{
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet("第一页");
		Row row = sheet.createRow(2);
		row.setHeightInPoints(30);
		
		CreateCell(workbook, row, (short)0, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_BOTTOM);
		CreateCell(workbook, row, (short)1, CellStyle.ALIGN_FILL, CellStyle.VERTICAL_TOP);
		CreateCell(workbook, row, (short)2, CellStyle.ALIGN_RIGHT, CellStyle.VERTICAL_JUSTIFY);
		
		FileOutputStream fileOutputStream = new FileOutputStream("D:\\gongzuobu.xls");
		workbook.write(fileOutputStream);
		fileOutputStream.close();
		
	}
	
	public static void CreateCell(Workbook workbook, Row row, short column ,short halign, short valign)
	{
		Cell cell = row.createCell(column);
		cell.setCellValue("hello");
		
		//单元格样式
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setAlignment(halign);
		cellStyle.setVerticalAlignment(valign);
		
		//某一行使用某一个单元格样式
		cell.setCellStyle(cellStyle);
	}

}
