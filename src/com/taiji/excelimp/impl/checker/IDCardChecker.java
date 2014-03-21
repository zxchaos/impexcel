package com.taiji.excelimp.impl.checker;

import java.util.regex.Pattern;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.taiji.excelimp.api.IRegExpChecker;
/**
 * 身份证号验证
 * @author zhangxin
 *
 */
public class IDCardChecker implements IRegExpChecker {
	private Logger logger = LoggerFactory.getLogger(IDCardChecker.class);
	@Override
	public boolean check(String cellValue) {
		logger.debug("---输入身份证号---"+cellValue+"+++");
		boolean result = false;
		Pattern idcp = Pattern.compile("(^\\d{15}$)|(^\\d{18}$)|(^\\d{17}(\\d|X|x)$)");
		if (idcp.matcher(cellValue).matches()) {
			result = true;
		}
		return result;
	}

}