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
public class XianName2CodeConverter extends StringSelectAbstractConverter {
	private Logger logger = LoggerFactory.getLogger(XianName2CodeConverter.class);
	@Override
	public String getCellValue(String cellValue) {
		logger.debug("---�����Ʊ���ת����---");
		logger.debug("---��Ԫ������---"+cellValue);
		String result = "";
		if (StringUtils.isNotBlank(cellValue) && !"��ֱ".equals(cellValue)) {
			HashMap<String, Long> xianMap = RegionUtil.getXianMap();
			Long converted = xianMap.get(cellValue);
			if (converted != null) {
				result = String.valueOf(converted);
			}
		}else if ("��ֱ".equals(cellValue)) {
			result = "-1";
		}
		logger.debug("---ת����ֵ---"+result);
		return result;
	}

}
