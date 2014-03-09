package com.taiji.excelimp.impl.converter;

/**
 * 线路等级转换器
 * @author zhangxin
 *
 */
public class XLDJConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		if ("一级网络".equals(cellValue)) {
			result = "1";
		}else if ("二级网络".equals(cellValue)) {
			result = "2";
		}else if ("三级网络".equals(cellValue)) {
			result = "3";
		}
		return result;
	}

}
