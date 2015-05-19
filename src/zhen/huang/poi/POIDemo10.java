package zhen.huang.poi;


import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

//×ÖÌå²âÊÔ
public class POIDemo10
{
	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException
	{
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet();
		Row row = sheet.createRow(0);
		Cell cell= row.createCell(0);
		
		CellStyle cellStyle = workbook.createCellStyle();
		Font font =  workbook.createFont();
		font.setItalic(true);
		font.setStrikeout(true);
		font.setBold(true);
		font.setFontHeightInPoints((short) 30);
		font.setFontName("Courier New");
		
		cell.setCellValue("Hello");
		cellStyle.setFont(font);
		cell.setCellStyle(cellStyle);
		
		FileOutputStream fileOutputStream = new FileOutputStream("D:\\¹¤×÷²¾4.xls");
		workbook.write(fileOutputStream);
		fileOutputStream.close();
	}

}
