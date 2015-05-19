package zhen.huang.poi;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


// 设置工作簿单元格边框样式 利用cellStyle.setBorderXXX 与 cellStyle.SetXXXBorDerColor()

public class POIDemo8
{
	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException
	{
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet("第一页");
		Row row = sheet.createRow(2);
		
		Cell cell2 = row.createCell(4);
		cell2.setCellValue("hello2");
		CellStyle cellStyle2 = workbook.createCellStyle();
		cellStyle2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		cellStyle2.setFillForegroundColor(IndexedColors.RED.getIndex());
		cell2.setCellStyle(cellStyle2);
		
		FileOutputStream fileOutputStream = new FileOutputStream("D:\\工作簿2.xls");
		workbook.write(fileOutputStream);
		fileOutputStream.close();
		
		
		
		
	}

}
