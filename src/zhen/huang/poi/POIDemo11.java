package zhen.huang.poi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

//���Զ�ȡ����д��Ԫ��
public class POIDemo11
{
	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException
	{
		InputStream inputStream =new FileInputStream("D:\\����.xls");
		POIFSFileSystem poifsFileSystem = new POIFSFileSystem(inputStream);
		Workbook workbook = new HSSFWorkbook(poifsFileSystem);
		
		Sheet sheet = workbook.getSheetAt(0);
		if(sheet == null)
		{
			return ;
		}
		
		Row row = sheet.getRow(0);
		if (row == null)
		{
			return ;
		}
		
		Cell cell = row.getCell(3);
		
		if (cell != null)
		{
			row.removeCell(cell);
		}
		
		FileOutputStream fileOutputStream = new FileOutputStream("D:\\����.xls");
		workbook.write(fileOutputStream);
		fileOutputStream.close();
		
		
		
	}
	

}
