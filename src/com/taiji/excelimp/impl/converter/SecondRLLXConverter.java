package com.taiji.excelimp.impl.converter;

import java.util.Arrays;

import org.apache.poi.ss.usermodel.Cell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.taiji.excelimp.api.IConverter;

/**
 * 燃料类型明细转换器
 * 
 * @author zhangxin
 * 
 */
public class SecondRLLXConverter implements IConverter {
	private Logger logger = LoggerFactory.getLogger(SecondRLLXConverter.class);
	@Override
	public String convert(Cell cell) throws Exception {
		String result = "";
		if (cell != null && Cell.CELL_TYPE_STRING == cell.getCellType()) {
			String cellValue = cell.getRichStringCellValue().getString();
			logger.debug("---燃料明细单元格值---"+cellValue);
			Cell rllxCell = cell.getRow().getCell(cell.getColumnIndex()-1);//获得燃料类型单元格
			if (rllxCell == null || Cell.CELL_TYPE_STRING != rllxCell.getCellType()) {
				return result;
			}
			String rllx = rllxCell.getRichStringCellValue().getString();
			logger.debug("---之前的燃料类型单元格值---"+rllx);
			if ("汽油".equals(rllx) && "单燃料-汽油".equals(cellValue)) {
				result = cellValue;
			} else if ("柴油".equals(rllx) && "单燃料-柴油".equals(cellValue)) {
				result = cellValue;
			} else if ("LPG".equals(rllx) && "单燃料-LPG".equals(cellValue)) {
				result = cellValue;
			} else if ("CNG".equals(rllx) && "单燃料-CNG".equals(cellValue)) {
				result = cellValue;
			} else if ("LNG".equals(rllx) && "单燃料-LNG".equals(cellValue)) {
				result = cellValue;
			} else if ("双燃料".equals(rllx)) {
				String [] doubleFuel = new String []{"汽油+LPG","汽油+CNG","汽油+LNG","柴油+LPG","柴油+CNG","柴油+LNG"};
				if (Arrays.asList(doubleFuel).contains(cellValue)) {
					result = cellValue;
				}
			} else if ("混合动力".equals(rllx)) {
				String [] mixFuel = new String [] {"汽油+电","柴油+电"};
				if (Arrays.asList(mixFuel).contains(cellValue)) {
					result = cellValue;
				}
			}
		}
		return result;
	}
	
	

}
