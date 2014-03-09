package com.taiji.excelimp;

import java.util.Properties;
import java.util.TimerTask;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.taiji.excelimp.core.AbstractImpExcel;
import com.taiji.excelimp.db.DBAccess;
import com.taiji.excelimp.factory.api.ImpExcelFactory;


/**
 * 轮询定时任务抽象类
 * @author zhangxin
 *
 */
public class ImportTask extends TimerTask {
	public static Logger logger = LoggerFactory.getLogger(ImportTask.class);
	private Properties impConfig = null;
	private DBAccess dBAccess = null;
	private ImpExcelFactory impExcelFactory;
	
	public ImpExcelFactory getImpExcelFactory() {
		return impExcelFactory;
	}

	public void setImpExcelFactory(ImpExcelFactory impExcelFactory) {
		this.impExcelFactory = impExcelFactory;
	}

	public DBAccess getDbAccess() {
		return dBAccess;
	}

	public void setDBAccess(DBAccess dbAccess) {
		this.dBAccess = dbAccess;
	}

	public Properties getImpConfig() {
		return impConfig;
	}

	public void setImpConfig(Properties impConfig) {
		this.impConfig = impConfig;
	}
	
	@Override
	public void run() {
		logger.info("+++轮询开始+++");
		try {
			AbstractImpExcel impExcel = impExcelFactory.getImpExcelInstance();
			impExcel.importExcel(this.getImpConfig(), this.getDbAccess());
		} catch (Exception e) {
			logger.error(e.getMessage(),e);
		}
		logger.info("+++结束轮询+++\n\n");
	}
}
