package com.taiji.excelimp.impl.converter;
/**
 * ��Ӫģʽ�����˵�ѡ��ת��
 * @author zhangxin
 *
 */
public class YYMSConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		if ("����".equals(cellValue) || "˫��".equals(cellValue)||"����".equals(cellValue)) {
			result = cellValue;
		}
		return result;
	}
}
