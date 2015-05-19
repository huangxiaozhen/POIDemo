package zhen.huang.poi;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

//测试自定义格式
public class POIDemo13
{
	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException
	{
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet();
		DataFormat dataFormat = wb.createDataFormat();
		Row row;
		Cell cell;
		CellStyle cellStyle;
		int rowNum = 0;

		row = sheet.createRow(rowNum++);
		cell = row.createCell(0);
		cell.setCellValue(111111.25);

		cellStyle = wb.createCellStyle();
		cellStyle.setDataFormat(dataFormat.getFormat("0.0"));
		cell.setCellStyle(cellStyle);

		row = sheet.createRow(rowNum++);
		cell = row.createCell(0);
		cell.setCellValue(111111.25);

		cellStyle = wb.createCellStyle();
		cellStyle.setDataFormat(dataFormat.getFormat("#,##0.000"));
		cell.setCellStyle(cellStyle);

		FileOutputStream fileOutputStream = new FileOutputStream(
				"D:\\POI的工作簿.xls");

		wb.write(fileOutputStream);

		fileOutputStream.close();

	}
}
