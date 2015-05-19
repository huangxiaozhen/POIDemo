package zhen.huang.poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class POIDemo2
{

	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException
	{
		Workbook workbook = new HSSFWorkbook();
		workbook.createSheet("第一个 sheet 页");
		workbook.createSheet("第二个 sheet 页");
		FileOutputStream fileOutputStream = new FileOutputStream(
				"D:\\POI的第一个sheet页.xls");

		workbook.write(fileOutputStream);

		fileOutputStream.close();
	}

}
