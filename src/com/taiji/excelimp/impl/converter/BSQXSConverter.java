package com.taiji.excelimp.impl.converter;

import java.util.Arrays;

/**
 * ��������ʽת����
 * @author zhangxin
 *
 */
public class BSQXSConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		cellValue = cellValue.replace("(", "").replace(")", "").replace("��", "").replace("��", "").replace("/", "");
		String [] bsqxsArray = new String [] {"�ֶ�������MT","�Զ�������AT","�ֶ��Զ�������","�޼�ʽ������"};
		if (Arrays.asList(bsqxsArray).contains(cellValue)) {
			result = cellValue;
			if ("�ֶ�������MT".equals(cellValue)) {
				result = "�ֶ�������(MT)";
			}else if ("�Զ�������AT".equals(cellValue)) {
				result = "�Զ�������(AT)";
			}else if ("�ֶ��Զ�������".equals(cellValue)) {
				result = "�ֶ�/�Զ�������";
			}else {
				result = cellValue;
			}
		}
		return result;
	}

}
