package com.taiji.excelimp.impl.converter;


/**
 * ������ɫת����
 * 
 * @author zhangxin
 * 
 */
public class CPYSConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		if ("��ɫ".equals(cellValue) || "��ɫ".equals(cellValue) || "����".equals(cellValue)) {
			result = cellValue;
		}
		return result;
	}
}
