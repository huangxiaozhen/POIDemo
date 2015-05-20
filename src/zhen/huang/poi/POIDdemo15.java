package zhen.huang.poi;

import java.io.ByteArrayOutputStream;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.zkoss.bind.annotation.Command;
import org.zkoss.zul.Filedownload;

import com.conf.object.Customer;
import com.conf.object.CustomerField;
import com.conf.object.CustomerFieldValue;

//客户导出模板 包含自定义字段 第一行表头 第二行 模板数据
//此程序为自己基于公司Saas 平台上写的，不能直接运行
public class POIDdemo15
{
	// 客户导入参考模板
	@Command
	public void uploadTemplete() throws Exception
	{
		Map<String, String> sheetHeaders = Customer.getCHColumns();
		Map<String, String> templateData = Customer.getTemplateData();
		ByteArrayOutputStream out = new ByteArrayOutputStream();

		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("客户导入模板页");
		Row row0 = sheet.createRow(0);
		Row row1 = sheet.createRow(1);
		int cellIndex = 0;

		for (String key : sheetHeaders.keySet())
		{
			row0.createCell(cellIndex).setCellValue(sheetHeaders.get(key));
			row1.createCell(cellIndex++).setCellValue(templateData.get(key));
		}

		// 加入自定义字段
		List<CustomerField> customerFieldList = customerFieldDao
				.findByTenantId(user.getTenant().getId());
		CustomerField customerField;

		CreationHelper creationHelper = workbook.getCreationHelper();
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(
				"yyyy-mm-dd hh:mm:ss"));
		for (int cellNum = 0; cellNum < customerFieldList.size(); cellNum++)
		{
			// 每一个自定义字段
			customerField = customerFieldList.get(cellNum);
			if (customerField.getType() == 0) // 文本框
			{
				if (customerField.isRequired())
				{
					row0.createCell(cellIndex).setCellValue(
							"*" + customerField.getName() + "（文本框）");
				} else
				{
					row0.createCell(cellIndex).setCellValue(
							customerField.getName() + "（文本框）");
				}
				row1.createCell(cellIndex++).setCellValue("这是测试文本数据");

			} else if (customerField.getType() == 4) // 时间框
			{
				if (customerField.isRequired())
				{
					row0.createCell(cellIndex).setCellValue(
							"*" + customerField.getName() + "(时间框）");
				} else
				{
					row0.createCell(cellIndex).setCellValue(
							customerField.getName() + "（时间框）");
				}
				row1.createCell(cellIndex).setCellStyle(cellStyle);
				row1.createCell(cellIndex++).setCellValue(new Date());

			} else if (customerField.getType() == 2) // 复选框
			{
				if (customerField.isRequired())
				{
					row0.createCell(cellIndex).setCellValue(
							"*" + customerField.getName() + "(复选框）");
				} else
				{
					row0.createCell(cellIndex).setCellValue(
							customerField.getName() + "（复选框）");
				}
				row1.createCell(cellIndex++).setCellValue(
						getCellValue(customerField, 2));

			} else if (customerField.getType() == 3) // 单选框
			{
				if (customerField.isRequired())
				{
					row0.createCell(cellIndex).setCellValue(
							"*" + customerField.getName() + "(单选框）");
				} else
				{
					row0.createCell(cellIndex).setCellValue(
							customerField.getName() + "（单选框）");
				}
				row1.createCell(cellIndex++).setCellValue(
						getCellValue(customerField, 3));

			} else if (customerField.getType() == 1) // 下拉框
			{
				if (customerField.isRequired())
				{
					row0.createCell(cellIndex).setCellValue(
							"*" + customerField.getName() + "(下拉框）");
				} else
				{
					row0.createCell(cellIndex).setCellValue(
							customerField.getName() + "（下拉框）");
				}
				row1.createCell(cellIndex++).setCellValue(
						getCellValue(customerField, 4));
			}
		}

		for (int i = 0; i < cellIndex; i++)
		{
			// 自适应宽度
			sheet.autoSizeColumn(i, true);
		}

		workbook.write(out);
		String filename = "客户导入模板.xls";
		Filedownload.save(out.toByteArray(), "application/octet-stream",
				filename);

	}

	// 得到复选框/单选框/下拉框的内容
	public String getCellValue(CustomerField customerField, int type)
	{
		List<CustomerFieldValue> customerFieldValueList;
		CustomerFieldValue customerFieldValue;
		customerFieldValueList = customerFieldValueDAO
				.findByFieldId(customerField.getId());
		String cellValue = "";
		if (type == 2)
		{
			for (CustomerFieldValue customerFieldValueTemp : customerFieldValueList)
			{
				cellValue = cellValue + customerFieldValueTemp.getFvValue()
						+ ",";
			}
			cellValue = cellValue.substring(0, cellValue.length() - 1);

		} else
		{
			customerFieldValue = customerFieldValueList.get(0);
			cellValue = customerFieldValue.getFvValue();
		}

		return cellValue;

	}

}
