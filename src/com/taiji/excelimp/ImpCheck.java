package com.taiji.excelimp;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Properties;
import java.util.Timer;

import org.apache.commons.lang3.StringUtils;
import org.omg.CORBA.PRIVATE_MEMBER;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.taiji.excelimp.db.DBAccess;
import com.taiji.excelimp.factory.api.ImpExcelFactory;
import com.taiji.excelimp.util.RegionUtil;

/**
 * 导入目录轮询查看
 * 
 * @author zhangxin
 * 
 */
public class ImpCheck {
	public static Logger logger = LoggerFactory.getLogger(ImpCheck.class);
	private static long period;
	private static long firstTimeDelay;
	private static Timer timer;
	private static Properties config;
	public static void main(String[] args) {
		
		String configFilePath = ImpCheck.class.getResource("/").getPath() + "impsysconfig.properties";
		try {
			initProp(configFilePath);
			
			String impFactoriesPropVal = config.getProperty("impExcelFactories");
			if (StringUtils.isBlank(impFactoriesPropVal)) {
				throw new Exception(configFilePath + "中没有配置impExcelFactories属性程序意外退出");
			}
			
			initRegionCache(new DBAccess(config.getProperty("dburl"), config.getProperty("username"),
					config.getProperty("password"), config.getProperty("driverClassName")));
			
			String[] factories = impFactoriesPropVal.split(":");
			for (int i = 0; i < factories.length; i++) {
				String factoryClass = factories[i];
				startTask(factoryClass);
			}

		} catch (Exception e) {
			logger.error(e.getMessage(), e);
		}

	}

	/**
	 * 初始化静态属性
	 * @param configFilePath
	 * @throws Exception
	 */
	private static void initProp(String configFilePath) throws Exception {
		timer = new Timer();
		logger.debug("---轮询系统配置文件路径---" + configFilePath);
		config = ImpCheck.readProperties(configFilePath);
		period = Long.valueOf(config.getProperty("period"));
		firstTimeDelay = Long.valueOf(config.getProperty("firstTimeDelay"));
		logger.debug("---period---" + period + "---firstTimeDelay---" + firstTimeDelay);
	}

	/**
	 * 启动定时任务
	 * @param factoryClass
	 * @throws InstantiationException
	 * @throws IllegalAccessException
	 * @throws ClassNotFoundException
	 */
	private static void startTask(String factoryClass) throws InstantiationException, IllegalAccessException,
			ClassNotFoundException {
		ImportTask importTask = new ImportTask();
		importTask.setImpConfig(config);
		importTask.setDBAccess(new DBAccess(config.getProperty("dburl"), config.getProperty("username"), config
				.getProperty("password"), config.getProperty("driverClassName")));
		ImpExcelFactory impExcelFactory = (ImpExcelFactory) Class.forName(factoryClass).newInstance();
		logger.debug("---导入Excel功能模块工厂对应的类---" + impExcelFactory.getClass().getName());
		importTask.setImpExcelFactory(impExcelFactory);
		timer.schedule(importTask, firstTimeDelay, period);
	}

	/**
	 * 初始化行者区划缓存
	 * 
	 * @param dbAccess
	 */
	private static void initRegionCache(DBAccess dbAccess) {
		String shiSQL = "select distinct(shi),shiname,shengname from t_sys_user_count t where t.shengname='新疆' and shiname is not null";
		HashMap<String, Long> shiMap = dbAccess.initHashMap(shiSQL);
		String xianSQL = "select xian,xianname from t_sys_user_count t where t.shengname='新疆' and xianname is not null";
		HashMap<String, Long> xianMap = dbAccess.initHashMap(xianSQL);
		RegionUtil.initShiXianMap(shiMap, xianMap);
		logger.debug("---缓存市县map完毕---");

	}

	/**
	 * 读取系统配置
	 * 
	 * @return
	 * @throws Exception
	 */
	public static Properties readProperties(String propPath) throws Exception {
		Properties sysConfig = new Properties();
		File propFile = new File(propPath);
		if (!propFile.exists()) {
			logger.error("文件:" + propFile.getAbsolutePath() + "不存在");
			throw new Exception("文件:" + propFile.getAbsolutePath() + "不存在");
		}
		FileInputStream fis = new FileInputStream(propFile);
		sysConfig.load(fis);
		return sysConfig;
	}

}
