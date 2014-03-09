package com.taiji.excelimp.impl.converter;

import java.util.ArrayList;
import java.util.Arrays;

/**
 * 车辆类型转换器
 * @author zhangxin
 *
 */
public class CLLXConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		String[] clxhArray = new String[] { "大型客车", "中型客车", "小型客车", "轿车", "大型卧铺客车", "中型卧铺客车", "大型双层客车", "大型铰接客车", "其他" };
		ArrayList<String>clxhList = new ArrayList<String>();
		if (Arrays.asList(clxhArray).contains(cellValue)) {
			result = cellValue;
		}
		return result;
	}

}
