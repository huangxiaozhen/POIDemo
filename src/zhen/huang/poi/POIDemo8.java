package zhen.huang.poi;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
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
		Cell cell = row.createCell(2);
		cell.setCellValue("hello");
		
		CellStyle cellStyle = workbook.createCellStyle();
		
		cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
		cellStyle.setBottomBorderColor(IndexedColors.RED.getIndex());
		
		cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
		cellStyle.setLeftBorderColor(IndexedColors.GREEN.getIndex());
		
		cellStyle.setBorderTop(CellStyle.BORDER_DASHED);
		cellStyle.setTopBorderColor(IndexedColors.RED.getIndex());
		
		cellStyle.setBorderRight(CellStyle.BORDER_THIN);
		cellStyle.setRightBorderColor(IndexedColors.RED.getIndex());
		
		cell.setCellStyle(cellStyle);
		
		FileOutputStream fileOutputStream = new FileOutputStream("D:\\������1.xls");
		workbook.write(fileOutputStream);
		fileOutputStream.close();
		
		
		
		
	}

}
