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
 * 单元格合并
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
		cell.setCellValue("这是一个单元格合并的测试");
		
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 1));
		FileOutputStream fileOutputStream  = new FileOutputStream("D:\\单元格测试.xls");
		workbook.write(fileOutputStream);
		fileOutputStream.close();
	}

}
