package zhen.huang.poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import zhen.huang.object.Customer;
import zhen.huang.utils.Utils;

//批量导出
public class POIDemo14
{
	@SuppressWarnings({"resource" })
	public static void main(String[] args) throws IOException
	{
		String headers[] = {"编号","姓名","年龄","职位"};
		
		Workbook workbook = new HSSFWorkbook();
		List<Customer> customerList = new ArrayList<Customer>();
		Customer customer1 = new Customer("1", "hunagzhen1", "test1", 24);
		Customer customer2 = new Customer("2", "huangzhen2", "test2", 24);
		Customer customer3 = new Customer("3", "huangzhen3", "test3", 24);
		Customer customer4 = new Customer("4", "huangzhen4", "test4", 24);
		Customer customer5 = new Customer("5", "huangzhen5", "test5", 24);
		
		customerList.add(customer1);
		customerList.add(customer2);
		customerList.add(customer3);
		customerList.add(customer4);
		customerList.add(customer5);
		
		workbook = Utils.excelExport(customerList, workbook, headers);
		
		
		FileOutputStream fileOutputStream = new FileOutputStream("D:\\customerList1.xls");
		FileOutputStream fileOutputStream1 = new FileOutputStream("D:\\customeTemplaterList.xls");
		workbook.write(fileOutputStream);
		Utils.excelExportByTemplate(customerList, "Template.xls").write(fileOutputStream1);
        fileOutputStream.close();		
        fileOutputStream1.close();

	}
}
