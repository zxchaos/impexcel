package com.taiji.excelimp.impl.converter;
/**
 * ��Ӫ��ʽת����
 * @author zhangxin
 *
 */
public class JYFSConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		if ("��Ӫ".equals(cellValue) || "����Ӫ".equals(cellValue)) {
			result = cellValue;
		}
		return result;
	}

}
