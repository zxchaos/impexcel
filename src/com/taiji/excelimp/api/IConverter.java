package com.taiji.excelimp.api;

import org.apache.poi.ss.usermodel.Cell;

/**
 * excel �е�Ԫ���ֶ�ת���ӿ�
 * @author zhangxin
 *
 */
public interface IConverter {

	public  String convert(Cell cell) throws Exception;
}
