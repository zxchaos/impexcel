package com.taiji.excelimp.impl.converter;
/**
 * 运营模式下拉菜单选项转换
 * @author zhangxin
 *
 */
public class YYMSConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		if ("单班".equals(cellValue) || "双班".equals(cellValue)||"其他".equals(cellValue)) {
			result = cellValue;
		}
		return result;
	}
}
