package com.taiji.excelimp.impl.converter;
/**
 * �Ƿ�ѡ��ת����
 * @author zxchaos
 *
 */
public class TrueFlaseConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		if ("��".equals(cellValue) || "��".equals(cellValue)) {
			result = cellValue;
		}
		return result;
	}

}
