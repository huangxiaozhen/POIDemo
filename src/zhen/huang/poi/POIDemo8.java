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


// ���ù�������Ԫ��߿���ʽ ����cellStyle.setBorderXXX �� cellStyle.SetXXXBorDerColor()

public class POIDemo8
{
	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException
	{
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet("��һҳ");
		Row row = sheet.createRow(2);
		
		Cell cell2 = row.createCell(4);
		cell2.setCellValue("hello2");
		CellStyle cellStyle2 = workbook.createCellStyle();
		cellStyle2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		cellStyle2.setFillForegroundColor(IndexedColors.RED.getIndex());
		cell2.setCellStyle(cellStyle2);
		
		FileOutputStream fileOutputStream = new FileOutputStream("D:\\������2.xls");
		workbook.write(fileOutputStream);
		fileOutputStream.close();
		
		
		
		
	}

}
