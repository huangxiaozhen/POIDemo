package zhen.huang.poi;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class POIDemo5
{
	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException
	{
		//得到一个输入流
		InputStream inputStream = new FileInputStream("D:\\名单.xls");

		//将输入流加到POI的文件系统中
		POIFSFileSystem fs = new POIFSFileSystem(inputStream);

		//创建一个工作簿,构造参数中加入POI的文件接口
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(fs);

		//创建sheet
		HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);

		if (hssfSheet == null)
		{
			return;
		}

		//按照 行/单元格 循环 sheet
		for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++)
		{
			HSSFRow hssfRow = hssfSheet.getRow(rowNum);
			if (hssfRow == null)
			{
				continue;
			}

			for (int cellNum = 0; cellNum <= hssfRow.getLastCellNum(); cellNum++)
			{
				HSSFCell hssfCell = hssfRow.getCell(cellNum);
				if (hssfCell == null)
				{
					continue;
				}
				System.out.println(getValue(hssfCell));
			}
			System.out.println("\n");
		}

	}

	public static String getValue(HSSFCell hssfCell)
	{
		if (hssfCell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN)
		{
			return String.valueOf(hssfCell.getBooleanCellValue());

		} else if (hssfCell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC)
		{
			return String.valueOf(hssfCell.getNumericCellValue());
		} else
		{
			return String.valueOf(hssfCell.getStringCellValue());
 
		}

	}

}
