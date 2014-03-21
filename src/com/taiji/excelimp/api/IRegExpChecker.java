package com.taiji.excelimp.api;

import org.apache.poi.ss.usermodel.Cell;

/**
 * 正则表达式检查接口
 * @author zhangxin
 *
 */
public interface IRegExpChecker {
	public boolean check(String cellValue);
}