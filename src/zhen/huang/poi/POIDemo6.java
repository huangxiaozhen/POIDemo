package zhen.huang.poi;

//提取 excel 文本操作
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class POIDemo6
{
	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException
	{
		//1. 得到文件的输入流
		InputStream inputStream = new FileInputStream("D:\\名单.xls");
		
		//2. 加入到 poi 文件系统中
		POIFSFileSystem poifsFileSystem =new POIFSFileSystem(inputStream);
		
		//3. 创建工作簿，把 poi 放到工作簿中
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(poifsFileSystem);
		
		//通过exce提取提取直接输出到控制台
		ExcelExtractor excelExtractor = new ExcelExtractor(hssfWorkbook);
		excelExtractor.setIncludeSheetNames(false);
		System.out.println(excelExtractor.getText());
		inputStream.close();
		
		
		
		
	}
	
	

}
