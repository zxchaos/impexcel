package com.taiji.excelimp.impl.converter;
/**
 * 是否选项转换器
 * @author zxchaos
 *
 */
public class TrueFlaseConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		if ("是".equals(cellValue) || "否".equals(cellValue)) {
			result = cellValue;
		}
		return result;
	}

}