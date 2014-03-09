package com.taiji.excelimp.api;

import org.apache.poi.ss.usermodel.Cell;

/**
 * excel 中单元格字段转换接口
 * @author zhangxin
 *
 */
public interface IConverter {

	public  String convert(Cell cell) throws Exception;
}
