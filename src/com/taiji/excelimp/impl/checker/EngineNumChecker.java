package com.taiji.excelimp.impl.checker;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.taiji.excelimp.api.IRegExpChecker;
/**
 * �������Ź����飺��������Ϊ6��14λ�ĺ��ֻ���ĸ�����ֵ����
 * @author zhangxin
 *
 */
public class EngineNumChecker implements IRegExpChecker {

	@Override
	public boolean check(String cellValue) {
		boolean result = false;
		String engineNumRegExp = "[\u4E00-\u9FA5a-zA-Z0-9]{6,14}";
		Pattern enPattern = Pattern.compile(engineNumRegExp);
		if (enPattern.matcher(cellValue).matches()) {
			result = true;
		}
		return result;
	}

}
