package com.taiji.excelimp.impl.converter;

import org.apache.poi.ss.usermodel.Cell;

import com.taiji.excelimp.api.IConverter;

/**
 * 燃油消耗数据来源转换器
 * 
 * @author zhangxin
 * 
 */
public class RYXHDataSourceConverter implements IConverter {

	@Override
	public String convert(Cell cell) throws Exception {
		String result = "";
		if (cell != null) {
			String cellValue = "";
			if (cell != null) {
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_BLANK:
					return result;
				case Cell.CELL_TYPE_STRING:
					cellValue = cell.getRichStringCellValue().getString();
					break;
				case Cell.CELL_TYPE_NUMERIC:
					cellValue = String.valueOf(cell.getNumericCellValue());
					break;
				default:
					return result;
				}
				// 燃油消耗数据来源：加油小票 台帐 加油加气卡 经验估算
				if ("加油小票".equals(cellValue) || "台帐".equals(cellValue) || "加油加气卡".equals(cellValue)
						|| "经验估算".equals(cellValue)) {
					result = cellValue;
				}
			}
		}
		return result;
	}
}