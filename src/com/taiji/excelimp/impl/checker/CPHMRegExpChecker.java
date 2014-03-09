package com.taiji.excelimp.impl.checker;

import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;

import com.taiji.excelimp.api.IRegExpChecker;
/**
 * ³µÅÆºÅÂë¼ìÑéÆ÷
 * @author zhangxin
 *
 */
public class CPHMRegExpChecker implements IRegExpChecker {

	@Override
	public boolean check(String cellValue) {
		boolean result = false;
		///^[a-zA-Z]{1}[a-zA-Z_0-9]{5}$/ 
		Pattern pattern = Pattern.compile("^ÐÂ[a-zA-Z]([a-zA-Z_0-9]{5})$");
		if (StringUtils.isNotBlank(cellValue)) {
			result = pattern.matcher(cellValue).matches();
		}
		return result;
	}

}
