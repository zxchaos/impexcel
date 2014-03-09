package com.taiji.excelimp.impl.converter;

import java.util.Arrays;

/**
 * 燃料类型转换器
 * @author zhangxin
 *
 */
public class RLLXConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String [] rllxArray = new String[]{"汽油", "柴油", "LPG","CNG", "LNG", "双燃料","混合动力"};
		String result = "";
		if (Arrays.asList(rllxArray).contains(cellValue)) {
			result = cellValue;
		}
		return result;
	}

}
