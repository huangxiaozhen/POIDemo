package zhen.huang.poi;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import zhen.huang.object.Customer;
import zhen.huang.object.Student;
import zhen.huang.utils.Utils;

//通用的导出模板
public class POIDemo15
{
	@SuppressWarnings({ "resource" })
	public static void main(String[] args) throws Exception
	{
		// 1. 客户导出
		String headers[] =
		{ "编号", "姓名", "年龄", "职位" };

		Workbook workbook = new HSSFWorkbook();
		List<Customer> customerList = new ArrayList<Customer>();
		Customer customer1 = new Customer("1", "hunagzhen1", "test1", 24);
		Customer customer2 = new Customer("2", "huangzhen2", "test2", 24);
		Customer customer3 = new Customer("3", "huangzhen3", "test3", 24);
		Customer customer4 = new Customer("4", "huangzhen4", "test4", 24);
		Customer customer5 = new Customer("5", "huangzhen5", "test5", 24);
		Customer customer6= new Customer("6", "huangzhen5", "test5", 24);

		customerList.add(customer1);
		customerList.add(customer2);
		customerList.add(customer3);
		customerList.add(customer4);
		customerList.add(customer5);
		customerList.add(customer6);
		
		workbook = Utils.export("Customer",customerList, workbook, headers, Customer.class);

		//2. 学生导出
		String studentHeaders[] =
		{ "学号", "姓名", "性别", "分数","身高" };

		Workbook workbookStudent = new HSSFWorkbook();
		List<Student> studentList = new ArrayList<Student>();
		Student student1 = new Student("1001", "黄小镇1", "男", 80, 171.5);
		Student student2 = new Student("1002", "黄小镇2", "男", 80, 171.5);
		Student student3 = new Student("1003", "黄小镇3", "男", 80, 171.5);
		Student student4 = new Student("1004", "黄小镇4", "男", 80, 171.5);
		Student student5 = new Student("1005", "黄小镇5", "男", 80, 171.5);
		Student student6 = new Student("1006", "黄小镇6", "男", 80, 171.5);
		Student student7 = new Student("1007", "黄小镇7", "男", 80, 171.5);
		Student student8 = new Student("1008", "黄小镇8", "男", 80, 171.5);
		Student student9 = new Student("1009", "黄小镇9", "男", 80, 171.5);

		studentList.add(student1);
		studentList.add(student2);
		studentList.add(student3);
		studentList.add(student4);
		studentList.add(student5);
		studentList.add(student6);
		studentList.add(student7);
		studentList.add(student8);
		studentList.add(student9);
		studentList.add(student9);
		studentList.add(student9);
		studentList.add(student9);
		
		workbookStudent = Utils.export("Student",studentList, workbookStudent, studentHeaders, Student.class);
		
		//3. 加入到文件输出流，进行输出
		FileOutputStream fileOutputStream = new FileOutputStream("D:\\customer.xls");
		workbook.write(fileOutputStream);
		
		fileOutputStream = new FileOutputStream("D:\\student.xls");
		workbookStudent.write(fileOutputStream);
		
		fileOutputStream.close();
		

	}

}
