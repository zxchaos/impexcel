package com.taiji.excelimp.impl.converter;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 使用性质转换器
 * @author zhangxin
 *
 */
public class SYXZConverter extends StringSelectAbstractConverter {
	private Logger logger = LoggerFactory.getLogger(SYXZConverter.class);
	@Override
	public String getCellValue(String cellValue) {
		logger.debug("---使用性质转换器----传入的值为"+cellValue+"---");
		String result = "";
		if ("非营运".equals(cellValue) || "公路客运".equals(cellValue) || "公交客运".equals(cellValue) || "出租客运".equals(cellValue)
				|| "旅游客运".equals(cellValue)||"租赁".equals(cellValue) || "营转非".equals(cellValue) || "出租转非".equals(cellValue) || "其他".equals(cellValue)) {
			result = cellValue;
		}
		logger.debug("---result---"+result);
		return result;
	}

}