package com.taiji.excelimp.impl.converter;

/**
 * ��·�ȼ�ת����
 * @author zhangxin
 *
 */
public class XLDJConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		if ("һ������".equals(cellValue)) {
			result = "1";
		}else if ("��������".equals(cellValue)) {
			result = "2";
		}else if ("��������".equals(cellValue)) {
			result = "3";
		}
		return result;
	}

}
