package com.taiji.excelimp.impl.converter;

/**
 * ��������ת����
 * @author zhangxin
 *
 */
public class BXLXConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		if ("��ͨ".equals(cellValue)) {
			result = "1";
		}else if ("ֱ��".equals(cellValue)) {
			result = "2";
		}
		return result;
	}
}
