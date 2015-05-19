package zhen.huang.poi;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/*
 * ��Ԫ��ϲ�
 */
public class POIDemo9
{
	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException
	{
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet();
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue("����һ����Ԫ��ϲ��Ĳ���");
		
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 1));
		FileOutputStream fileOutputStream  = new FileOutputStream("D:\\��Ԫ�����.xls");
		workbook.write(fileOutputStream);
		fileOutputStream.close();
	}

}
