package com.taiji.excelimp.impl.converter;


/**
 * 变更情况转换器
 * @author zhangxin
 *
 */
public class BGQKConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		if ("新增".equals(cellValue) || "过户转入".equals(cellValue)) {
			result = cellValue;
		}
		return result;
	}

	
}
