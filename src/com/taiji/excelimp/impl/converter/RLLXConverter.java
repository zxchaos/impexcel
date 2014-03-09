package com.taiji.excelimp.impl.converter;

import java.util.Arrays;

/**
 * ȼ������ת����
 * @author zhangxin
 *
 */
public class RLLXConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String [] rllxArray = new String[]{"����", "����", "LPG","CNG", "LNG", "˫ȼ��","��϶���"};
		String result = "";
		if (Arrays.asList(rllxArray).contains(cellValue)) {
			result = cellValue;
		}
		return result;
	}

}
