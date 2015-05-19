package zhen.huang.poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class POIDemo1  
{
	public static void main(String[] args) throws IOException
	{

		Workbook wb = new HSSFWorkbook();
		
		FileOutputStream fileOutputStream = new FileOutputStream("D:\\POI的第一个工作簿.xls");
		
		wb.write(fileOutputStream);
		
		fileOutputStream.close();
		
	}
	
	

}
