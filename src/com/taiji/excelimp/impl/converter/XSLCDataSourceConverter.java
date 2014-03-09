package com.taiji.excelimp.impl.converter;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;

import com.taiji.excelimp.api.IConverter;

/**
 * 行驶里程数据来源转换器
 * 
 * @author zhangxin
 * 
 */
public class XSLCDataSourceConverter implements IConverter {

	@Override
	public String convert(Cell cell) throws Exception {
		String result = "";
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
			
			// 行驶里程数据来源：GPS 台帐 电子路单 行驶记录仪 经验估算
			if ("GPS".equals(cellValue) || "台帐".equals(cellValue) || "电子路单".equals(cellValue)
					|| "行驶记录仪".equals(cellValue) || "经验估算".equals(cellValue)) {
				result = cellValue;
			}
		}
		return result;
	}

}
