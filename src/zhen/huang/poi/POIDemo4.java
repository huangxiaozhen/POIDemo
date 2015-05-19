package zhen.huang.poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

//对于时间的支持
public class POIDemo4
{
	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException
	{
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet("这是一页");
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		
		cell.setCellValue(new Date());
		
		//设置单元格样式
		CreationHelper creationHelper = workbook.getCreationHelper();
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
		
		Cell cell2 = row.createCell(1);
		cell2.setCellValue(new Date());
		cell2.setCellStyle(cellStyle);
		
		Cell cell3 = row.createCell(2);
		cell3.setCellValue(Calendar.getInstance());
		cell3.setCellStyle(cellStyle);
		
		
		FileOutputStream fileOutputStream = new FileOutputStream("D:\\工作簿.xls");
		workbook.write(fileOutputStream);
		fileOutputStream.close();
	}
	

}
