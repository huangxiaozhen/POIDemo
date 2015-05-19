package zhen.huang.poi;

//��ȡ excel �ı�����
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
		//1. �õ��ļ���������
		InputStream inputStream = new FileInputStream("D:\\����.xls");
		
		//2. ���뵽 poi �ļ�ϵͳ��
		POIFSFileSystem poifsFileSystem =new POIFSFileSystem(inputStream);
		
		//3. �������������� poi �ŵ���������
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(poifsFileSystem);
		
		//ͨ��exce��ȡ��ȡֱ�����������̨
		ExcelExtractor excelExtractor = new ExcelExtractor(hssfWorkbook);
		excelExtractor.setIncludeSheetNames(false);
		System.out.println(excelExtractor.getText());
		inputStream.close();
		
		
		
		
	}
	
	

}
