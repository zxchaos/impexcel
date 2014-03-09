package com.taiji.excelimp.impl.checker;

import java.util.regex.Pattern;

import com.taiji.excelimp.api.IRegExpChecker;
/**
 * 身份证号验证
 * @author zhangxin
 *
 */
public class IDCardChecker implements IRegExpChecker {

	@Override
	public boolean check(String cellValue) {
		boolean result = false;
		Pattern idcp = Pattern.compile("(^\\d{15}$)|(^\\d{18}$)|(^\\d{17}(\\d|X|x)$)");
		if (idcp.matcher(cellValue).matches()) {
			result = true;
		}
		return result;
	}

}
