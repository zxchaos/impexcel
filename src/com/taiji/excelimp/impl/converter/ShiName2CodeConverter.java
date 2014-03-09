package com.taiji.excelimp.impl.converter;

import java.util.HashMap;

import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.taiji.excelimp.util.RegionUtil;
/**
 * �����Ʊ���ת����
 * @author zhangxin
 *
 */
public class ShiName2CodeConverter extends StringSelectAbstractConverter {
	private static Logger logger = LoggerFactory.getLogger(ShiName2CodeConverter.class);
	@Override
	public String getCellValue(String cellValue) {
		logger.debug("---�����Ʊ���ת����---");
		logger.debug("---��Ԫ��ֵ---"+cellValue);
		String result = "";
		if (StringUtils.isNotBlank(cellValue)) {
			HashMap<String, Long> shiMap = RegionUtil.getShiMap();
			Long converted = shiMap.get(cellValue);
			if (converted != null) {
				result = String.valueOf(shiMap.get(cellValue));
			}
		}
		logger.debug("---ת����ֵ---"+result);
		return result;
	}

}
