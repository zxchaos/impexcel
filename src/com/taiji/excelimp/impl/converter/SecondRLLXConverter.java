package com.taiji.excelimp.impl.converter;

import java.util.Arrays;

import org.apache.poi.ss.usermodel.Cell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.taiji.excelimp.api.IConverter;

/**
 * ȼ��������ϸת����
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
			logger.debug("---ȼ����ϸ��Ԫ��ֵ---"+cellValue);
			Cell rllxCell = cell.getRow().getCell(cell.getColumnIndex()-1);//���ȼ�����͵�Ԫ��
			if (rllxCell == null || Cell.CELL_TYPE_STRING != rllxCell.getCellType()) {
				return result;
			}
			String rllx = rllxCell.getRichStringCellValue().getString();
			logger.debug("---֮ǰ��ȼ�����͵�Ԫ��ֵ---"+rllx);
			if ("����".equals(rllx) && "��ȼ��-����".equals(cellValue)) {
				result = cellValue;
			} else if ("����".equals(rllx) && "��ȼ��-����".equals(cellValue)) {
				result = cellValue;
			} else if ("LPG".equals(rllx) && "��ȼ��-LPG".equals(cellValue)) {
				result = cellValue;
			} else if ("CNG".equals(rllx) && "��ȼ��-CNG".equals(cellValue)) {
				result = cellValue;
			} else if ("LNG".equals(rllx) && "��ȼ��-LNG".equals(cellValue)) {
				result = cellValue;
			} else if ("˫ȼ��".equals(rllx)) {
				String [] doubleFuel = new String []{"����+LPG","����+CNG","����+LNG","����+LPG","����+CNG","����+LNG"};
				if (Arrays.asList(doubleFuel).contains(cellValue)) {
					result = cellValue;
				}
			} else if ("��϶���".equals(rllx)) {
				String [] mixFuel = new String [] {"����+��","����+��"};
				if (Arrays.asList(mixFuel).contains(cellValue)) {
					result = cellValue;
				}
			}
		}
		return result;
	}
	
	

}
