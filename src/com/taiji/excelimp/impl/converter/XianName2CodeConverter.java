package com.taiji.excelimp.impl.converter;

import java.util.HashMap;

import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.taiji.excelimp.util.RegionUtil;

/**
 * 县名称编码转换器
 * @author zhangxin
 *
 */
public class XianName2CodeConverter extends StringSelectAbstractConverter {
	private Logger logger = LoggerFactory.getLogger(XianName2CodeConverter.class);
	@Override
	public String getCellValue(String cellValue) {
		logger.debug("---县名称编码转换器---");
		logger.debug("---单元格名称---"+cellValue);
		String result = "";
		if (StringUtils.isNotBlank(cellValue) && !"市直".equals(cellValue)) {
			HashMap<String, Long> xianMap = RegionUtil.getXianMap();
			Long converted = xianMap.get(cellValue);
			if (converted != null) {
				result = String.valueOf(converted);
			}
		}else if ("市直".equals(cellValue)) {
			result = "-1";
		}
		logger.debug("---转换后值---"+result);
		return result;
	}

}
