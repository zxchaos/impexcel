package com.taiji.excelimp.factory;

import com.taiji.excelimp.core.AbstractImpExcel;
import com.taiji.excelimp.core.RegisterBatchImpExcel;
import com.taiji.excelimp.factory.api.ImpExcelFactory;

public class RegisterBatchImpExcelFactory implements ImpExcelFactory {

	@Override
	public AbstractImpExcel getImpExcelInstance() throws Exception {
		RegisterBatchImpExcel registerBatchImpExcel = new RegisterBatchImpExcel();
		registerBatchImpExcel.setType("2");
		return registerBatchImpExcel;
	}

}
