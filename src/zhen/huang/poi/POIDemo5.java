package zhen.huang.poi;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class POIDemo5
{
	public static void main(String[] args) throws IOException
	{
		InputStream inputStream = new FileInputStream("D:\\Ãûµ¥.xls");

		POIFSFileSystem fs = new POIFSFileSystem(inputStream);

		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(fs);

		HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);

		if (hssfSheet == null)
		{
			return;
		}

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
