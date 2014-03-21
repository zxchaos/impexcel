package com.taiji.excelimp.impl.checker;

import java.util.regex.Pattern;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.taiji.excelimp.api.IRegExpChecker;
import com.taiji.excelimp.util.ExcelConstants;

/**
 * 日期正则表达式检验
 * @author zhangxin
 *
 */
public class DateRegExpChecker implements IRegExpChecker {
	private Logger logger = LoggerFactory.getLogger(DateRegExpChecker.class);
	@Override
	public boolean check(String cellValue) {
		logger.debug("---日期输入值---"+cellValue);
		boolean result = false;
		Pattern pattern = Pattern.compile(ExcelConstants.REGEXP_DATE);
		result = pattern.matcher(cellValue).matches();
		logger.debug("---验证结果---"+result);
		return result;
	}
	
}