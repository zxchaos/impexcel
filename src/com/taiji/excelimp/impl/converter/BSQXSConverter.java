package com.taiji.excelimp.impl.converter;

import java.util.Arrays;

/**
 * 变速器型式转换器
 * @author zhangxin
 *
 */
public class BSQXSConverter extends StringSelectAbstractConverter {

	@Override
	public String getCellValue(String cellValue) {
		String result = "";
		cellValue = cellValue.replace("(", "").replace(")", "").replace("（", "").replace("）", "").replace("/", "");
		String [] bsqxsArray = new String [] {"手动变速器MT","自动变速器AT","手动自动变速器","无级式变速器"};
		if (Arrays.asList(bsqxsArray).contains(cellValue)) {
			result = cellValue;
			if ("手动变速器MT".equals(cellValue)) {
				result = "手动变速器(MT)";
			}else if ("自动变速器AT".equals(cellValue)) {
				result = "自动变速器(AT)";
			}else if ("手动自动变速器".equals(cellValue)) {
				result = "手动/自动变速器";
			}else {
				result = cellValue;
			}
		}
		return result;
	}

}