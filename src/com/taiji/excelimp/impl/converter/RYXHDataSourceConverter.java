package com.taiji.excelimp.impl.converter;

import org.apache.poi.ss.usermodel.Cell;

import com.taiji.excelimp.api.IConverter;

/**
 * ȼ������������Դת����
 * 
 * @author zhangxin
 * 
 */
public class RYXHDataSourceConverter implements IConverter {

	@Override
	public String convert(Cell cell) throws Exception {
		String result = "";
		if (cell != null) {
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
				// ȼ������������Դ������СƱ ̨�� ���ͼ����� �������
				if ("����СƱ".equals(cellValue) || "̨��".equals(cellValue) || "���ͼ�����".equals(cellValue)
						|| "�������".equals(cellValue)) {
					result = cellValue;
				}
			}
		}
		return result;
	}
}
