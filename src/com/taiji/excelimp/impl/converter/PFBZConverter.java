package com.taiji.excelimp.impl.converter;

import java.util.Arrays;

import org.apache.commons.lang3.ArrayUtils;

/**
 * 排放标准转换器
 * @author zhangxin
 *
 */
public class PFBZConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		String [] pfbzArray = new String[]{"国Ⅳ以上", "国Ⅳ", "国Ⅲ", "国Ⅱ", "国Ⅱ以下"};
		if (Arrays.asList(pfbzArray).contains(cellValue)) {
			result = cellValue;
		}
		return result;
	}

}
