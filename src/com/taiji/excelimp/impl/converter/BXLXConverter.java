package com.taiji.excelimp.impl.converter;

/**
 * 班线类型转换器
 * @author zhangxin
 *
 */
public class BXLXConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		if ("普通".equals(cellValue)) {
			result = "1";
		}else if ("直达".equals(cellValue)) {
			result = "2";
		}
		return result;
	}
}
