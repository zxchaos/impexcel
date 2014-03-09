package com.taiji.excelimp.factory;

import com.taiji.excelimp.core.AbstractImpExcel;
import com.taiji.excelimp.core.NckyxlBatchImpExcel;
import com.taiji.excelimp.factory.api.ImpExcelFactory;

public class NckyxlglImpExcelFactory implements ImpExcelFactory {

	@Override
	public AbstractImpExcel getImpExcelInstance() throws Exception {
		NckyxlBatchImpExcel nckyxlglImpExcel = new NckyxlBatchImpExcel();
		nckyxlglImpExcel.setType("3");
		return nckyxlglImpExcel;
	}

}
