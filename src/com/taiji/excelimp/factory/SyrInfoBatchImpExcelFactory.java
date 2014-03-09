package com.taiji.excelimp.factory;

import com.taiji.excelimp.core.AbstractImpExcel;
import com.taiji.excelimp.core.SyrInfoBatchImpExcel;
import com.taiji.excelimp.factory.api.ImpExcelFactory;

public class SyrInfoBatchImpExcelFactory implements ImpExcelFactory {

	@Override
	public AbstractImpExcel getImpExcelInstance() throws Exception {
		SyrInfoBatchImpExcel syrInfoBatchImpExcel = new SyrInfoBatchImpExcel();
		syrInfoBatchImpExcel.setType("4");
		return syrInfoBatchImpExcel;
	}

}
