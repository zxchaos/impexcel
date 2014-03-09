package com.taiji.excelimp.impl.checker;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.taiji.excelimp.api.IRegExpChecker;
/**
 * ��ؿ��������������ؿ�������Ҫ���ڵ���0��С�ڵ���999
 * @author zhangxin
 *
 */
public class EDZKChecker implements IRegExpChecker {
	private Logger logger = LoggerFactory.getLogger(EDZKChecker.class);
	@Override
	public boolean check(String cellValue) {
		logger.debug("---��ؿ����������---����ֵ---"+cellValue);
		boolean result = false;
		String intValue = "";
		if (cellValue.contains(".")) {
			intValue = cellValue.substring(0, cellValue.lastIndexOf("."));
		}else {
			intValue = cellValue;
		}
		logger.debug("---ת��ֵ���ֵ---"+intValue);
		if (Integer.valueOf(intValue)>=0 && Integer.valueOf(intValue)<=999) {
			result = true;
		}
		return result;
	}

}
