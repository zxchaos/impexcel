package com.taiji.excelimp.factory;

import com.taiji.excelimp.core.AbstractImpExcel;
import com.taiji.excelimp.core.DataBatchImpExcel;
import com.taiji.excelimp.factory.api.ImpExcelFactory;

public class DataBatchImpExcelFactory implements ImpExcelFactory {
	@Override
	public AbstractImpExcel getImpExcelInstance() throws Exception {
		DataBatchImpExcel dataBatchImpExcel = new DataBatchImpExcel();
		dataBatchImpExcel.setType("1");
		return dataBatchImpExcel;
	}

}
