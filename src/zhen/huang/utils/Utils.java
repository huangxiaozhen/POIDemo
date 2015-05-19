package zhen.huang.utils;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import zhen.huang.object.Customer;

//如何将这个写成一个通用的方法，如果要写成通用的方法必须要用到反射的ji's
public class Utils
{
	public static Workbook excelExport(List<Customer> list, Workbook workbook,
			String[] headers)
	{
		int rowIndex = 0;
		Sheet sheet = workbook.createSheet();
		Row row = sheet.createRow(rowIndex++);

		for (int cellNum = 0; cellNum < headers.length; cellNum++)
		{
			row.createCell(cellNum).setCellValue(headers[cellNum]);
		}

		Iterator<Customer> iterator = list.iterator();
		while (iterator.hasNext())
		{
			int cellNum = 0;
			row = sheet.createRow(rowIndex++);
			Customer customer = iterator.next();
			row.createCell(cellNum++).setCellValue(customer.getId());
			row.createCell(cellNum++).setCellValue(customer.getName());
			row.createCell(cellNum++).setCellValue(customer.getAge());
			row.createCell(cellNum++).setCellValue(customer.getJob());
		}

		return workbook;

	}

	public static Workbook excelExportByTemplate(List<Customer> list,
			String templateName) throws IOException
	{
		InputStream inputStream = Utils.class
				.getResourceAsStream("/zhen/huang/template/" + templateName);

		POIFSFileSystem poifsFileSystem = new POIFSFileSystem(inputStream);
		Workbook workbook = new HSSFWorkbook(poifsFileSystem);
		Sheet sheet = workbook.getSheetAt(0);

		int rowIndex = 1;
		Iterator<Customer> iterator = list.iterator();
		while (iterator.hasNext())
		{
			Row row = sheet.createRow(rowIndex++);
			int cellNum = 0;
			Customer customer = iterator.next();
			row.createCell(cellNum++).setCellValue(customer.getId());
			row.createCell(cellNum++).setCellValue(customer.getName());
			row.createCell(cellNum++).setCellValue(customer.getAge());
			row.createCell(cellNum++).setCellValue(customer.getJob());
		}

		return workbook;

	}

	// 通用的方法
	public static Workbook export(String sheetName,List<?> list, Workbook workbook,
			String[] headers, Class<?> clz) throws Exception
	{
		int rowIndex = 0;

		//设置excel样式
		workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet(sheetName);
		
		
		Row row = sheet.createRow(rowIndex++);

		// 写入 excel 表格的头部
		for (int cellNum = 0; cellNum < headers.length; cellNum++)
		{
			row.createCell(cellNum).setCellValue(headers[cellNum]);
			
		}

		// 写入 excel表格的主题部分

		Field[] fields = clz.getDeclaredFields(); // 得到对应的Class对应的属性。
													// 也就是类的属性值有哪些

		// 遍历 list 中的对象
		for (int i = 0; i < list.size(); i++)
		{
			// list里面的每一个对象对应的就是一行
			row = sheet.createRow(rowIndex++);

			// set 单元格的值 遍历对象的每个属性
			for (int j = 0; j < fields.length; j++)
			{
				// 对private 属性需要设置
				fields[j].setAccessible(true);
				row.createCell(j).setCellValue(fields[j].get(list.get(i)).toString());
			}

		}

		// 设置单元格宽度
		for (int i = 0; i < headers.length; i++)
		{
			//自适应宽度
			sheet.autoSizeColumn(i, true);
			if (sheet.getColumnWidth(i) < (headers[i].getBytes().length) * 400)
			{
				sheet.setColumnWidth(i,(headers[i].getBytes().length) * 400);
			}
		}

		return workbook;

	}

}
