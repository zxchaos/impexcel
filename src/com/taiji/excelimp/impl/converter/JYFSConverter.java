package com.taiji.excelimp.impl.converter;
/**
 * 经营方式转换器
 * @author zhangxin
 *
 */
public class JYFSConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		if ("自营".equals(cellValue) || "非自营".equals(cellValue)) {
			result = cellValue;
		}
		return result;
	}

}