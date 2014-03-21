package com.taiji.excelimp.impl.converter;

import java.util.HashMap;

import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.taiji.excelimp.util.RegionUtil;
/**
 * 市名称编码转换器
 * @author zhangxin
 *
 */
public class ShiName2CodeConverter extends StringSelectAbstractConverter {
	private static Logger logger = LoggerFactory.getLogger(ShiName2CodeConverter.class);
	@Override
	public String getCellValue(String cellValue) {
		logger.debug("---市名称编码转换器---");
		logger.debug("---单元格值---"+cellValue);
		String result = "";
		if (StringUtils.isNotBlank(cellValue)) {
			HashMap<String, Long> shiMap = RegionUtil.getShiMap();
			Long converted = shiMap.get(cellValue);
			if (converted != null) {
				result = String.valueOf(shiMap.get(cellValue));
			}
		}
		logger.debug("---转换后值---"+result);
		return result;
	}

}