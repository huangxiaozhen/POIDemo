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
		workbook.createSheet("��һ�� sheet ҳ");
		workbook.createSheet("�ڶ��� sheet ҳ");
		FileOutputStream fileOutputStream = new FileOutputStream(
				"D:\\POI�ĵ�һ��sheetҳ.xls");

		workbook.write(fileOutputStream);

		fileOutputStream.close();
	}

}
