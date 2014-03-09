package com.taiji.excelimp.impl.checker;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.taiji.excelimp.api.IRegExpChecker;
/**
 * 额定载客人数检查器：额定载客人数需要大于等于0，小于等于999
 * @author zhangxin
 *
 */
public class EDZKChecker implements IRegExpChecker {
	private Logger logger = LoggerFactory.getLogger(EDZKChecker.class);
	@Override
	public boolean check(String cellValue) {
		logger.debug("---额定载客人数检查器---传入值---"+cellValue);
		boolean result = false;
		String intValue = "";
		if (cellValue.contains(".")) {
			intValue = cellValue.substring(0, cellValue.lastIndexOf("."));
		}else {
			intValue = cellValue;
		}
		logger.debug("---转换值后的值---"+intValue);
		if (Integer.valueOf(intValue)>=0 && Integer.valueOf(intValue)<=999) {
			result = true;
		}
		return result;
	}

}
