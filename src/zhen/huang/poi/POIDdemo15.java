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

//�ͻ�����ģ�� �����Զ����ֶ� ��һ�б�ͷ �ڶ��� ģ������
//�˳���Ϊ�Լ����ڹ�˾Saas ƽ̨��д�ģ�����ֱ������
public class POIDdemo15
{
	// �ͻ�����ο�ģ��
	@Command
	public void uploadTemplete() throws Exception
	{
		Map<String, String> sheetHeaders = Customer.getCHColumns();
		Map<String, String> templateData = Customer.getTemplateData();
		ByteArrayOutputStream out = new ByteArrayOutputStream();

		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("�ͻ�����ģ��ҳ");
		Row row0 = sheet.createRow(0);
		Row row1 = sheet.createRow(1);
		int cellIndex = 0;

		for (String key : sheetHeaders.keySet())
		{
			row0.createCell(cellIndex).setCellValue(sheetHeaders.get(key));
			row1.createCell(cellIndex++).setCellValue(templateData.get(key));
		}

		// �����Զ����ֶ�
		List<CustomerField> customerFieldList = customerFieldDao
				.findByTenantId(user.getTenant().getId());
		CustomerField customerField;

		CreationHelper creationHelper = workbook.getCreationHelper();
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(
				"yyyy-mm-dd hh:mm:ss"));
		for (int cellNum = 0; cellNum < customerFieldList.size(); cellNum++)
		{
			// ÿһ���Զ����ֶ�
			customerField = customerFieldList.get(cellNum);
			if (customerField.getType() == 0) // �ı���
			{
				if (customerField.isRequired())
				{
					row0.createCell(cellIndex).setCellValue(
							"*" + customerField.getName() + "���ı���");
				} else
				{
					row0.createCell(cellIndex).setCellValue(
							customerField.getName() + "���ı���");
				}
				row1.createCell(cellIndex++).setCellValue("���ǲ����ı�����");

			} else if (customerField.getType() == 4) // ʱ���
			{
				if (customerField.isRequired())
				{
					row0.createCell(cellIndex).setCellValue(
							"*" + customerField.getName() + "(ʱ���");
				} else
				{
					row0.createCell(cellIndex).setCellValue(
							customerField.getName() + "��ʱ���");
				}
				row1.createCell(cellIndex).setCellStyle(cellStyle);
				row1.createCell(cellIndex++).setCellValue(new Date());

			} else if (customerField.getType() == 2) // ��ѡ��
			{
				if (customerField.isRequired())
				{
					row0.createCell(cellIndex).setCellValue(
							"*" + customerField.getName() + "(��ѡ��");
				} else
				{
					row0.createCell(cellIndex).setCellValue(
							customerField.getName() + "����ѡ��");
				}
				row1.createCell(cellIndex++).setCellValue(
						getCellValue(customerField, 2));

			} else if (customerField.getType() == 3) // ��ѡ��
			{
				if (customerField.isRequired())
				{
					row0.createCell(cellIndex).setCellValue(
							"*" + customerField.getName() + "(��ѡ��");
				} else
				{
					row0.createCell(cellIndex).setCellValue(
							customerField.getName() + "����ѡ��");
				}
				row1.createCell(cellIndex++).setCellValue(
						getCellValue(customerField, 3));

			} else if (customerField.getType() == 1) // ������
			{
				if (customerField.isRequired())
				{
					row0.createCell(cellIndex).setCellValue(
							"*" + customerField.getName() + "(������");
				} else
				{
					row0.createCell(cellIndex).setCellValue(
							customerField.getName() + "��������");
				}
				row1.createCell(cellIndex++).setCellValue(
						getCellValue(customerField, 4));
			}
		}

		for (int i = 0; i < cellIndex; i++)
		{
			// ����Ӧ���
			sheet.autoSizeColumn(i, true);
		}

		workbook.write(out);
		String filename = "�ͻ�����ģ��.xls";
		Filedownload.save(out.toByteArray(), "application/octet-stream",
				filename);

	}

	// �õ���ѡ��/��ѡ��/�����������
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
