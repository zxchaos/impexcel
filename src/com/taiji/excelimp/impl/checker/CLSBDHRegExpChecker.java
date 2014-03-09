package com.taiji.excelimp.impl.checker;

import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;

import com.taiji.excelimp.api.IRegExpChecker;
/**
 * 车辆识别代号正则验证器
 * @author zhangxin
 *
 */
public class CLSBDHRegExpChecker implements IRegExpChecker {

	@Override
	public boolean check(String cellValue) {
		boolean result = false;
		if (StringUtils.isNotBlank(cellValue)) {
			Pattern pattern = Pattern.compile("[A-HJ-NPR-Za-hj-npr-z0-9]{17}");//共17位。不会包含 I、O、Q三个英文字母
			if (pattern.matcher(cellValue).matches()) {
				result = true;
			}
		}
		return result;
	}

}
