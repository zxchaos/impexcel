package com.taiji.excelimp.impl.converter;


/**
 * ������ת����
 * @author zhangxin
 *
 */
public class BGQKConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		if ("����".equals(cellValue) || "����ת��".equals(cellValue)) {
			result = cellValue;
		}
		return result;
	}

	
}
