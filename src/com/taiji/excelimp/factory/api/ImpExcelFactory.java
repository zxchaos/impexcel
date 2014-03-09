package com.taiji.excelimp.factory.api;

import com.taiji.excelimp.core.AbstractImpExcel;

public  interface ImpExcelFactory {
	public  AbstractImpExcel getImpExcelInstance() throws Exception;
}
