package zhen.huang.poi;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

//���Ի���
public class POIDemo12
{
	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException
	{
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet();
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue("���Ի��� \n �ɹ���");
		
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setWrapText(true);
		cell.setCellStyle(cellStyle);
		
		row.setHeightInPoints(2 * sheet.getDefaultRowHeightInPoints());
		sheet.autoSizeColumn(2);
		
		
		
		FileOutputStream fileOutputStream = new FileOutputStream("D:\\POI�ĵ�һ��������.xls");
		
		wb.write(fileOutputStream);
		
		fileOutputStream.close();

	}
}
