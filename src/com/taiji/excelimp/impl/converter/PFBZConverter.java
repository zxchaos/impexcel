package com.taiji.excelimp.impl.converter;

import java.util.Arrays;

import org.apache.commons.lang3.ArrayUtils;

/**
 * �ŷű�׼ת����
 * @author zhangxin
 *
 */
public class PFBZConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		String [] pfbzArray = new String[]{"��������", "����", "����", "����", "��������"};
		if (Arrays.asList(pfbzArray).contains(cellValue)) {
			result = cellValue;
		}
		return result;
	}

}
