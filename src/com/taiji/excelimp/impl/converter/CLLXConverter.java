package com.taiji.excelimp.impl.converter;

import java.util.ArrayList;
import java.util.Arrays;

/**
 * ��������ת����
 * @author zhangxin
 *
 */
public class CLLXConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		String[] clxhArray = new String[] { "���Ϳͳ�", "���Ϳͳ�", "С�Ϳͳ�", "�γ�", "�������̿ͳ�", "�������̿ͳ�", "����˫��ͳ�", "���ͽ½ӿͳ�", "����" };
		ArrayList<String>clxhList = new ArrayList<String>();
		if (Arrays.asList(clxhArray).contains(cellValue)) {
			result = cellValue;
		}
		return result;
	}

}
