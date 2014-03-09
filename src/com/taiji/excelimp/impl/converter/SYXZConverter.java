package com.taiji.excelimp.impl.converter;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * ʹ������ת����
 * @author zhangxin
 *
 */
public class SYXZConverter extends StringSelectAbstractConverter {
	private Logger logger = LoggerFactory.getLogger(SYXZConverter.class);
	@Override
	public String getCellValue(String cellValue) {
		logger.debug("---ʹ������ת����----�����ֵΪ"+cellValue+"---");
		String result = "";
		if ("��Ӫ��".equals(cellValue) || "��·����".equals(cellValue) || "��������".equals(cellValue) || "�������".equals(cellValue)
				|| "���ο���".equals(cellValue)||"����".equals(cellValue) || "Ӫת��".equals(cellValue) || "����ת��".equals(cellValue) || "����".equals(cellValue)) {
			result = cellValue;
		}
		logger.debug("---result---"+result);
		return result;
	}

}
