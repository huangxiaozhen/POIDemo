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
 * ���������ͻ��̶��ֶ��Լ��Զ����ֶγ���
 * 
 * �˳����ǲ��Գ����е�һ���֣�����ֱ������
 * 
 * �����Ѿ���ԭ�еĳ������ݲ���ͨ��
 * 
 *
 */

public class POIDemo17
{
	@Command
	public void downloadTest() throws Exception
	{
		// ���Ե����ͻ��̶��ֶ��Լ��ͻ����Զ����ֶ�
		/**
		 * �������������������� 1. �ͻ��Ĺ̶��ֶ� 2. �ͻ����Զ����ֶ� �̶��ֶδ洢�� t_customer ���У��Զ����ֶδ洢��
		 * ��������
		 * 
		 * t_customer_field id���Զ����ֶα�ʾ�� name���Զ����ֶ����֣� type(�Զ����ֶ����� ���Ͱ���
		 * �ı������ڣ���ѡ�򣬸�ѡ��������) required �Ƿ���� tenantId �Զ����ֶ������⻧
		 * 
		 * t_customer_field_value �Զ����ֶο�ѡ��ֵ id fieldId t_customer_field ���е� id
		 * fv_key �� 0 ��ʼ ���縴ѡ��ȶ��кü���ֵ fv_value ��Ӧ�� key ����Ӧ��ֵ
		 * 
		 * t_customer_field_corr �ͻ����Զ����ֶε�ֵ id customerId fieldId value
		 * ĳ���ͻ���Ӧ��ĳ���Զ������͵�ֵ ���� ������Ӧ���Ա��ֵ ��������
		 * 
		 * 
		 * �̶��ֶ�ֻ����һ������Ҫ�ģ��Զ�����ֶε�ֵ��Ҫȫ������ ��ͷ�������� ��Ҫ�����Ĺ̶��ֶε�ֵ + �Զ����ֶε�ֵ ��������� �ͻ�
		 */

		// ��Ҫ�����Ĺ̶��ֶ� Map �б� ��Key �ǿͻ��д��ڵ��ֶΣ�Value ��Ҫ�� excel ���ͷ����ʾ��
		Map<String, String> sheetHeaders = Customer.getCHColumns(); 
																	
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		
		//��ʼ����������
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("�ͻ���Ϣ");
		HSSFRow row = sheet.createRow(0);
		row.setHeightInPoints(20);

		// ��Ԫ����ʽ����
		HSSFCellStyle cellStyle = workbook.createCellStyle();
		HSSFFont font = workbook.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		HSSFCellStyle headStyle = workbook.createCellStyle();
		// ˮƽ�������þ��У���ֱ���򱣳�Ĭ��
		cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
		headStyle.setAlignment(CellStyle.ALIGN_CENTER);
		headStyle.setFont(font);

		//���÷���õ��������� Class ������������� 
		Field[] fields = Customer.class.getDeclaredFields();
		List<Field> fieldsTemp = new ArrayList<Field>();
		// �ж� Customer���������� ����Щ��������Ҫ�����ģ������б��� sheetHeaders ��
		for (Field field : fields)
		{
			if (sheetHeaders != null
					&& sheetHeaders.containsKey(field.getName()))
			{
				fieldsTemp.add(field);
			}
		}

		// Ϊ���Ժ���Ҫ������Щ�ͻ����ֶ���
		// List ת�������� Ϊʲôд�� new Field[0] �ο� CollectionDemo �е� ListToArray ����
		fields = fieldsTemp.toArray(new Field[0]);

		// �����̶��ֶεı�ͷ
		for (int cellNum = 0; cellNum < fields.length; cellNum++)
		{
			HSSFCell cell = row.createCell(cellNum);
			cell.setCellValue(sheetHeaders.get(fields[cellNum].getName()));
			cell.setCellStyle(headStyle);
			sheet.autoSizeColumn(cellNum, true);
		}

		// �����̶��ֶεĿͻ���Ϣ
		List<Customer> customerList = customerDao.getAll(); // ���ݲ�ѯ�����õ���Ҫ�����Ŀͻ��б�

		for (int rowNum = 0; rowNum < customerList.size(); rowNum++)
		{
			// һ���ͻ���Ӧһ��
			row = sheet.createRow(rowNum + 1);
			row.setHeightInPoints(20);

			// ���ñ�ͷ���Զ�Ӧ��ֵ
			for (int cellNum = 0; cellNum < fields.length; cellNum++)
			{
				fields[cellNum].setAccessible(true);
				HSSFCell cell = row.createCell(cellNum);

				String cellValue = fields[cellNum]
						.get(customerList.get(rowNum)) == null ? ""
						: fields[cellNum].get(customerList.get(rowNum))
								.toString();
				// ���ñ�ͷ��Ԫ���ֵ����ʽ
				cell.setCellValue(cellValue);
				cell.setCellStyle(cellStyle);
			}

		}

		// ���ù̶��ֶεĿ��
		Set<String> keySet = sheetHeaders.keySet();
		Iterator<String> it = keySet.iterator();
		for (int i = 0; i < fields.length; i++)
		{
			String temp = sheetHeaders.get(it.next());
			sheet.autoSizeColumn(i, true);

			System.out.println("��" + i + "�еĿ�ȣ�" + sheet.getColumnWidth(i));
			System.out.println("��" + i + "���ַ��Ŀ�ȣ�" + temp.getBytes().length);
			System.out.println("�ַ����Ŀ��" + "��".getBytes().length);
			System.out.println("�ַ����Ŀ��" + "a".getBytes().length);
			System.out.println("�ַ����Ŀ��" + "* ".getBytes().length);

			if (sheet.getColumnWidth(i) < (temp.getBytes().length) * 300)
			{
				sheet.setColumnWidth(i, (temp.getBytes().length) * 300);
			}

			System.out.println("��" + i + "������Ŀ�ȣ�" + sheet.getColumnWidth(i));

		}

		// �����Զ����ֶεı�ͷ
		// 1. �Զ����ͷ�� List
		List<CustomerField> customerFieldList = customerFieldDao
				.findByTenantId(user.getTenant().getId());
		// 2. ��� List �е� Name ����ֵ
		for (int cellNum = 0; cellNum < customerFieldList.size(); cellNum++)
		{
			row = sheet.getRow(0);
			HSSFCell cell = row.createCell(cellNum + sheetHeaders.size());
			String cellValue;
			// ����Ǳ����ֶ� ǰ����� *����ʶ
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

		// �����Զ����ֶε�ֵ
		// �ͻ��Զ����ֶε�ֵ ���� customerId �Լ� fieldId �õ�
		// �������ǿͻ�������
		for (int rowNum = 0; rowNum < customerList.size(); rowNum++)
		{
			// �õ�ÿһ���ͻ���Ӧ����
			row = sheet.getRow(rowNum + 1);
			Customer customer = customerList.get(rowNum);
			// ѭ�������Ԫ���ֵ
			for (int cellNum = 0; cellNum < customerFieldList.size(); cellNum++)
			{
				// �õ�ÿһ���Զ����ֶ� ��ʼ�д����еĹ̶��ֶ��п�ʼ
				HSSFCell cell = row.createCell(cellNum + sheetHeaders.size());
				String cellValue = "";
				CustomerField customerField = customerFieldList.get(cellNum);
						.findByCustomerIdAndFieldId(customer.getId(),
								customerField.getId());

				// �ж�����һ�����ͣ���ͬ�����͵õ�ֵ�ķ�ʽ��ͬ
				int fieldType = customerField.getType();

				if (fieldType == 0) // �ı��� �����Ϊ�գ�ֱ�ӰѶ�Ӧ��ֵд����Ӧ�ĵ�Ԫ����
				{
					if (!Utils.isEmpty(customerFieldCorr))
					{
						cellValue = customerFieldCorr.getValue();
					}
					

				} else if (fieldType == 4) // ʱ��� ע�����ʱ���ʽ
				{
					if (!Utils.isEmpty(customerFieldCorr))
					{
						cellValue = customerFieldCorr.getValue();
					}

				} else if (fieldType == 1) // ������ ֻ��һ��ֵ
				{
					cellValue = getCellValue(fieldType, customerField,
							customerFieldCorr);

				} else if (fieldType == 2) // ��ѡ�� ���ܶ��֮,��,����
				{
					cellValue = getCellValue(fieldType, customerField,
							customerFieldCorr);

				} else if (fieldType == 3) // ��ѡ�� ֻ��һ��ֻ
				{
					cellValue = getCellValue(fieldType, customerField,
							customerFieldCorr);

				}
				
				//�����Զ����ֶ��е�ֵ
				cell.setCellValue(cellValue);
				cell.setCellStyle(cellStyle);
			}
		}

		// ���� �Զ����ֶεĵ�����ʽ
		int cellLength = sheet.getRow(0).getLastCellNum();
		for (int cellNum = sheetHeaders.size(); cellNum < cellLength; cellNum++)
		{
			row = sheet.getRow(0);
			// �õ���ͷ��ĳ�е�������
			String cellValue = row.getCell(cellNum).getStringCellValue();
			sheet.autoSizeColumn(cellNum, true);

			// ������������ĳ�еĿ��
			if (sheet.getColumnWidth(cellNum) < (cellValue.getBytes().length) * 300)
			{
				sheet.setColumnWidth(cellNum,
						(cellValue.getBytes().length) * 300);
			}

		}

		// ��������������ͨ�����������
		workbook.write(out);
		String filename = "�ͻ�����.xls";
		Filedownload.save(out.toByteArray(), "application/octet-stream",
				filename);
	}

	public String getCellValue(int fieldType, CustomerField customerField,
			CustomerFieldCorr customerFieldCorr)
	{
		String cellValue = "";
		if (fieldType != 2) // ��ѡ�������� ����ʽ��ͬ
		{
			if (!Utils.isEmpty(customerFieldCorr))
			{
				List<CustomerFieldValue> customerFieldValueList = customerFieldValueDAO
						.findByFieldId(customerField.getId());

				// �ж� ��ѡ������������е����־���Ĵ���ֵ
				for (CustomerFieldValue customerFieldValue : customerFieldValueList)
				{
					if (customerFieldValue.getFvKey() == Integer
							.parseInt(customerFieldCorr.getValue()))
					{
						cellValue = customerFieldValue.getFvValue(); // �õ�
																		// ĳ���ͻ���Ӧ��ĳ���ֶε�
																		// ����������ľ���ֵ
					}
				}
			}

		} else
		// ��ѡ�� ���ֵ ������ԱȽϸ���
		{
			if (!Utils.isEmpty(customerFieldCorr))
			{
				String[] filedValueKey = customerFieldCorr.getValue()
						.split(",");

				// �ж�ÿһ�� Key ����Ӧ���ַ�����ֵ
				List<CustomerFieldValue> customerFieldValueList = customerFieldValueDAO
						.findByFieldId(customerField.getId());
				// �� key ����ѭ��
				for (int fieldValueKeyNum = 0; fieldValueKeyNum < filedValueKey.length; fieldValueKeyNum++)
				{
					// �ж� ĳ�� fieldValue ��ֵ�Ƿ���� keyֵ
					for (CustomerFieldValue customerFieldValue : customerFieldValueList)
					{
						if (customerFieldValue.getFvKey() == Integer
								.parseInt(filedValueKey[fieldValueKeyNum]))
						{
							// �����Ӧ�� key ��ͬ����Ѷ�Ӧ�� value ��ֵ��ֵ����Ԫ����
							cellValue = cellValue
									+ customerFieldValue.getFvValue() + "��";
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
