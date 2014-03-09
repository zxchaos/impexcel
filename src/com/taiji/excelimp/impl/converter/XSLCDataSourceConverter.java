package com.taiji.excelimp.impl.converter;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;

import com.taiji.excelimp.api.IConverter;

/**
 * ��ʻ���������Դת����
 * 
 * @author zhangxin
 * 
 */
public class XSLCDataSourceConverter implements IConverter {

	@Override
	public String convert(Cell cell) throws Exception {
		String result = "";
		String cellValue = "";
		if (cell != null) {
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_BLANK:
				return result;
			case Cell.CELL_TYPE_STRING:
				cellValue = cell.getRichStringCellValue().getString();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				cellValue = String.valueOf(cell.getNumericCellValue());
				break;
			default:
				return result;
			}
			
			// ��ʻ���������Դ��GPS ̨�� ����·�� ��ʻ��¼�� �������
			if ("GPS".equals(cellValue) || "̨��".equals(cellValue) || "����·��".equals(cellValue)
					|| "��ʻ��¼��".equals(cellValue) || "�������".equals(cellValue)) {
				result = cellValue;
			}
		}
		return result;
	}

}
