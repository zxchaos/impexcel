package com.taiji.excelimp.impl.converter;


/**
 * 车牌颜色转换器
 * 
 * @author zhangxin
 * 
 */
public class CPYSConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		if ("蓝色".equals(cellValue) || "黄色".equals(cellValue) || "其他".equals(cellValue)) {
			result = cellValue;
		}
		return result;
	}
}
