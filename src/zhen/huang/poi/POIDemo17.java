package zhen.huang.poi;

import java.io.ByteArrayOutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.zkoss.bind.annotation.Command;

/**
 * 
 * 批量导出客户固定字段以及自定义字段程序
 * 
 * 此程序是测试程序中的一部分，不能直接运行
 * 
 * 程序已经在原有的程序内容测试通过
 * 
 *
 */

public class POIDemo17
{
	@Command
	public void downloadTest() throws Exception
	{
		// 测试导出客户固定字段以及客户的自定义字段
		/**
		 * 表格数据由两个部分组成 1. 客户的固定字段 2. 客户的自定义字段 固定字段存储在 t_customer 表中，自定义字段存储在
		 * 三个表中
		 * 
		 * t_customer_field id（自定义字段表示） name（自定义字段名字） type(自定义字段类型 类型包括
		 * 文本框，日期，单选框，复选框，下拉框) required 是否必填 tenantId 自定义字段所属租户
		 * 
		 * t_customer_field_value 自定义字段可选的值 id fieldId t_customer_field 表中的 id
		 * fv_key 从 0 开始 例如复选框等都有好几个值 fv_value 相应的 key 所对应的值
		 * 
		 * t_customer_field_corr 客户的自定义字段的值 id customerId fieldId value
		 * 某个客户对应的某个自定义类型的值 例如 张三对应的性别的值 这里是男
		 * 
		 * 
		 * 固定字段只导出一部分需要的，自定义的字段的值需要全部导出 表头的数据由 需要导出的固定字段的值 + 自定义字段的值 表格数据由 客户
		 */

		// 需要导出的固定字段 Map 列表 ，Key 是客户中存在的字段，Value 是要在 excel 表格头部显示的
		Map<String, String> sheetHeaders = Customer.getCHColumns(); 
																	
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		
		//开始创建工作簿
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("客户信息");
		HSSFRow row = sheet.createRow(0);
		row.setHeightInPoints(20);

		// 单元格样式设置
		HSSFCellStyle cellStyle = workbook.createCellStyle();
		HSSFFont font = workbook.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		HSSFCellStyle headStyle = workbook.createCellStyle();
		// 水平方向设置居中，垂直方向保持默认
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		headStyle.setAlignment(CellStyle.ALIGN_CENTER);
		headStyle.setFont(font);

		//利用反射得到传进来的 Class 对象的所有属性 
		Field[] fields = Customer.class.getDeclaredFields();
		List<Field> fieldsTemp = new ArrayList<Field>();
		// 判断 Customer所有属性中 有哪些属性是需要导出的，导出列表在 sheetHeaders 中
		for (Field field : fields)
		{
			if (sheetHeaders != null
					&& sheetHeaders.containsKey(field.getName()))
			{
				fieldsTemp.add(field);
			}
		}

		// 为了以后需要导出哪些客户的字段用
		// List 转换成数组 为什么写成 new Field[0] 参考 CollectionDemo 中的 ListToArray 方法
		fields = fieldsTemp.toArray(new Field[0]);

		// 导出固定字段的表头
		for (int cellNum = 0; cellNum < fields.length; cellNum++)
		{
			HSSFCell cell = row.createCell(cellNum);
			cell.setCellValue(sheetHeaders.get(fields[cellNum].getName()));
			cell.setCellStyle(headStyle);
			sheet.autoSizeColumn(cellNum, true);
		}

		// 导出固定字段的客户信息
		List<Customer> customerList = customerDao.getAll(); // 根据查询条件得到需要导出的客户列表

		for (int rowNum = 0; rowNum < customerList.size(); rowNum++)
		{
			// 一个客户对应一行
			row = sheet.createRow(rowNum + 1);
			row.setHeightInPoints(20);

			// 设置表头属性对应的值
			for (int cellNum = 0; cellNum < fields.length; cellNum++)
			{
				fields[cellNum].setAccessible(true);
				HSSFCell cell = row.createCell(cellNum);

				String cellValue = fields[cellNum]
						.get(customerList.get(rowNum)) == null ? ""
						: fields[cellNum].get(customerList.get(rowNum))
								.toString();
				// 设置表头单元格的值和样式
				cell.setCellValue(cellValue);
				cell.setCellStyle(cellStyle);
			}

		}

		// 设置固定字段的宽度
		Set<String> keySet = sheetHeaders.keySet();
		Iterator<String> it = keySet.iterator();
		for (int i = 0; i < fields.length; i++)
		{
			String temp = sheetHeaders.get(it.next());
			sheet.autoSizeColumn(i, true);

			System.out.println("第" + i + "行的宽度：" + sheet.getColumnWidth(i));
			System.out.println("第" + i + "行字符的宽度：" + temp.getBytes().length);
			System.out.println("字符串的宽度" + "黄".getBytes().length);
			System.out.println("字符串的宽度" + "a".getBytes().length);
			System.out.println("字符串的宽度" + "* ".getBytes().length);

			if (sheet.getColumnWidth(i) < (temp.getBytes().length) * 300)
			{
				sheet.setColumnWidth(i, (temp.getBytes().length) * 300);
			}

			System.out.println("第" + i + "调整后的宽度：" + sheet.getColumnWidth(i));

		}

		// 导出自定义字段的表头
		// 1. 自定义表头的 List
		List<CustomerField> customerFieldList = customerFieldDao
				.findByTenantId(user.getTenant().getId());
		// 2. 输出 List 中的 Name 属性值
		for (int cellNum = 0; cellNum < customerFieldList.size(); cellNum++)
		{
			row = sheet.getRow(0);
			HSSFCell cell = row.createCell(cellNum + sheetHeaders.size());
			String cellValue;
			// 如果是必填字段 前面加上 *　标识
			if (customerFieldList.get(cellNum).isRequired())
			{
				cellValue = "* " + customerFieldList.get(cellNum).getName();
			} else
			{
				cellValue = customerFieldList.get(cellNum).getName();
			}
			cell.setCellValue(cellValue);
			cell.setCellStyle(headStyle);
		}

		// 导出自定义字段的值
		// 客户自定义字段的值 根据 customerId 以及 fieldId 得到
		// 行数还是客户的行数
		for (int rowNum = 0; rowNum < customerList.size(); rowNum++)
		{
			// 得到每一个客户对应的行
			row = sheet.getRow(rowNum + 1);
			Customer customer = customerList.get(rowNum);
			// 循环输出单元格的值
			for (int cellNum = 0; cellNum < customerFieldList.size(); cellNum++)
			{
				// 得到每一个自定义字段 起始列从所有的固定字段列开始
				HSSFCell cell = row.createCell(cellNum + sheetHeaders.size());
				String cellValue = "";
				CustomerField customerField = customerFieldList.get(cellNum);
						.findByCustomerIdAndFieldId(customer.getId(),
								customerField.getId());

				// 判断是哪一种类型，不同的类型得到值的方式不同
				int fieldType = customerField.getType();

				if (fieldType == 0) // 文本框 如果不为空，直接把对应的值写到对应的单元格中
				{
					if (!Utils.isEmpty(customerFieldCorr))
					{
						cellValue = customerFieldCorr.getValue();
					}
					

				} else if (fieldType == 4) // 时间框 注意输出时间格式
				{
					if (!Utils.isEmpty(customerFieldCorr))
					{
						cellValue = customerFieldCorr.getValue();
					}

				} else if (fieldType == 1) // 下拉框 只有一个值
				{
					cellValue = getCellValue(fieldType, customerField,
							customerFieldCorr);

				} else if (fieldType == 2) // 复选框 可能多个之,用,隔开
				{
					cellValue = getCellValue(fieldType, customerField,
							customerFieldCorr);

				} else if (fieldType == 3) // 单选框 只有一个只
				{
					cellValue = getCellValue(fieldType, customerField,
							customerFieldCorr);

				}
				
				//设置自定义字段列的值
				cell.setCellValue(cellValue);
				cell.setCellStyle(cellStyle);
			}
		}

		// 设置 自定义字段的导出样式
		int cellLength = sheet.getRow(0).getLastCellNum();
		for (int cellNum = sheetHeaders.size(); cellNum < cellLength; cellNum++)
		{
			row = sheet.getRow(0);
			// 得到表头的某列的中文名
			String cellValue = row.getCell(cellNum).getStringCellValue();
			sheet.autoSizeColumn(cellNum, true);

			// 根据条件设置某列的宽度
			if (sheet.getColumnWidth(cellNum) < (cellValue.getBytes().length) * 300)
			{
				sheet.setColumnWidth(cellNum,
						(cellValue.getBytes().length) * 300);
			}

		}

		// 将工作簿的内容通过浏览器下载
		workbook.write(out);
		String filename = "客户测试.xls";
		Filedownload.save(out.toByteArray(), "application/octet-stream",
				filename);
	}

	public String getCellValue(int fieldType, CustomerField customerField,
			CustomerFieldCorr customerFieldCorr)
	{
		String cellValue = "";
		if (fieldType != 2) // 单选框，下拉框 处理方式相同
		{
			if (!Utils.isEmpty(customerFieldCorr))
			{
				List<CustomerFieldValue> customerFieldValueList = customerFieldValueDAO
						.findByFieldId(customerField.getId());

				// 判断 单选框或者下拉框中的数字具体的代表值
				for (CustomerFieldValue customerFieldValue : customerFieldValueList)
				{
					if (customerFieldValue.getFvKey() == Integer
							.parseInt(customerFieldCorr.getValue()))
					{
						cellValue = customerFieldValue.getFvValue(); // 得到
																		// 某个客户对应的某个字段的
																		// 整数所代表的具体值
					}
				}
			}

		} else
		// 复选框 多个值 处理相对比较复杂
		{
			if (!Utils.isEmpty(customerFieldCorr))
			{
				String[] filedValueKey = customerFieldCorr.getValue()
						.split(",");

				// 判断每一个 Key 所对应的字符串的值
				List<CustomerFieldValue> customerFieldValueList = customerFieldValueDAO
						.findByFieldId(customerField.getId());
				// 对 key 进行循环
				for (int fieldValueKeyNum = 0; fieldValueKeyNum < filedValueKey.length; fieldValueKeyNum++)
				{
					// 判断 某个 fieldValue 的值是否等于 key值
					for (CustomerFieldValue customerFieldValue : customerFieldValueList)
					{
						if (customerFieldValue.getFvKey() == Integer
								.parseInt(filedValueKey[fieldValueKeyNum]))
						{
							// 如果对应的 key 相同，则把对应的 value 的值赋值到单元格中
							cellValue = cellValue
									+ customerFieldValue.getFvValue() + "，";
						}

					}
				}

				if (!Utils.isEmpty(cellValue))
				{
					cellValue = cellValue.substring(0, cellValue.length() - 1);
				}
			}
		}
		return cellValue;
	}

}
