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

//��ν����д��һ��ͨ�õķ��������Ҫд��ͨ�õķ�������Ҫ�õ������ji's
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

	// ͨ�õķ���
	public static Workbook export(String sheetName,List<?> list, Workbook workbook,
			String[] headers, Class<?> clz) throws Exception
	{
		int rowIndex = 0;

		//����excel��ʽ
		workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet(sheetName);
		
		
		Row row = sheet.createRow(rowIndex++);

		// д�� excel ����ͷ��
		for (int cellNum = 0; cellNum < headers.length; cellNum++)
		{
			row.createCell(cellNum).setCellValue(headers[cellNum]);
			
		}

		// д�� excel�������ⲿ��

		Field[] fields = clz.getDeclaredFields(); // �õ���Ӧ��Class��Ӧ�����ԡ�
													// Ҳ�����������ֵ����Щ

		// ���� list �еĶ���
		for (int i = 0; i < list.size(); i++)
		{
			// list�����ÿһ�������Ӧ�ľ���һ��
			row = sheet.createRow(rowIndex++);

			// set ��Ԫ���ֵ ���������ÿ������
			for (int j = 0; j < fields.length; j++)
			{
				// ��private ������Ҫ����
				fields[j].setAccessible(true);
				row.createCell(j).setCellValue(fields[j].get(list.get(i)).toString());
			}

		}

		// ���õ�Ԫ����
		for (int i = 0; i < headers.length; i++)
		{
			//����Ӧ���
			sheet.autoSizeColumn(i, true);
			if (sheet.getColumnWidth(i) < (headers[i].getBytes().length) * 400)
			{
				sheet.setColumnWidth(i,(headers[i].getBytes().length) * 400);
			}
		}

		return workbook;

	}

}
