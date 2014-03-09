package com.taiji.excelimp.impl.converter;

import org.apache.poi.ss.usermodel.Cell;

import com.taiji.excelimp.api.IConverter;
/**
 * ��excel�е�����ѡ���в������ϵͳ���������ת������excel�������ö��ֵ��Ϊ����ֵ����ͨ��ת����ģ��
 * @author zhangxin
 *
 */
public abstract class StringSelectAbstractConverter implements IConverter {

	@Override
	public String convert(Cell cell) throws Exception {
		String result = "";
		if (cell != null) {
			int cellType = cell.getCellType();
			if (Cell.CELL_TYPE_STRING == cellType) {
				String cellValue = cell.getRichStringCellValue().getString();
				result = getCellValue(cellValue);
			}
		}
		return result;
	}
	public abstract String getCellValue(String cellValue);
}
