package com.taiji.excelimp.impl.converter;

import org.apache.poi.ss.usermodel.Cell;

import com.taiji.excelimp.api.IConverter;
/**
 * 对excel中的下拉选项中不需根据系统中配置码表转换（即excel中所填的枚举值即为存入值）的通用转换器模板
 * @author zhangxin
 *
 */
public abstract class StringSelectAbstractConverter implements IConverter {

	@Override
	public String convert(Cell cell) throws Exception {
		String result = "";
		if (cell != null) {
			int cellType = cell.getCellType();
			if (Cell.CELL_TYPE_STRING == cellType) {
				String cellValue = cell.getRichStringCellValue().getString();
				result = getCellValue(cellValue);
			}
		}
		return result;
	}
	public abstract String getCellValue(String cellValue);
}