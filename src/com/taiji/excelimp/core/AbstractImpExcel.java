package com.taiji.excelimp.core;

import java.io.File;
import java.io.IOException;
import java.sql.Connection;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.dom4j.Document;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.taiji.excelimp.db.DBAccess;
import com.taiji.excelimp.util.ExcelConstants;
import com.taiji.excelimp.util.ExcelImportUtil;

public abstract class AbstractImpExcel {
	public static Logger logger = LoggerFactory.getLogger(AbstractImpExcel.class);
	/**模块类别：目前type取四个值，1：数据批量导入，2：车辆注册信息批量导入，3：农村客运线路批量导入，4：受益人信息批量导入*/
	private String type;

	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
	}

	/**
	 * 导入excelDir目录下的excel文件
	 * 
	 * @param excelDir
	 */
	public abstract void importExcel(Properties sysConfig, DBAccess dbAccess) throws Exception;

	/**
	 * @param impDir
	 * @return impDir下的文件 若该目录下无文件则返回null
	 * @throws Exception
	 */
	public File[] checkImpDir(String impDir) throws Exception {
		File excelDir = new File(impDir);
		if (!excelDir.exists()) {
			// 导入目录不存在则自动创建
			excelDir.mkdirs();
		} else if (!excelDir.isDirectory()) {
			logger.error("路径：" + impDir + "不是目录");
			throw new Exception("路径：" + impDir + "不是目录");
		}
		logger.info("+++轮询目录+++" + excelDir);
		// 选择excel 2003 与 2007的文件
		File[] excelFiles = ExcelImportUtil.getDirFiles(excelDir);
		if (excelFiles.length == 0) {
			logger.info("+++目录" + excelDir.getAbsolutePath() + "中无文件+++");
			return null;
		}
		logger.info("+++目录中的Excel文件个数+++" + excelFiles.length);
		return excelFiles;
	}

	/**
	 * 将错误信息写入日志
	 * 
	 * @param dbAccess
	 * @param resultMap
	 * @param infoFieldsMap 要插入日志的字段以及值的map。key：为字段名称，value：字段值
	 * @param type 模块类型：1：数据批量录入；2：车辆注册信息批量导入
	 *            
	 */
	public void insertImpInfo(DBAccess dbAccess, Map<String, String> resultMap, Map<String, String> infoFieldsMap,
			boolean isSuccess,String type) throws Exception {
		logger.debug("---将导入结果信息写入数据库---");
		long dwid = Long.valueOf(infoFieldsMap.get("dwid"));
		String hylb = infoFieldsMap.get("hylb");
		String sj = getNowDateString("yyyy年MM月dd日HH时mm分");
		String info = "";
		if (isSuccess) {
			info = "您导入的信息成功";
		} else {
			info = resultMap.get(ExcelConstants.MSG_KEY);
		}
		logger.info("+++导入结果信息+++\n" + info);
		if (StringUtils.isNotBlank(info) && info.length()>2000) {//若错误信息长度超过2000则自动截断为2000个字
			logger.info("+++错误信息超过2000个字符+++");
			info = info.substring(0,2001);
			logger.debug("---截断超过长度的错误信息后---"+info);
		}
//		dbAccess.insertErrorInfo(dwid, hylb, sj, info, type);
		logger.debug("---导入结果信息写入数据库完成---");
	}

	/**
	 * 检查文件是否为系统提供的模板
	 * 
	 * @param resultMap
	 *            检查结果
	 * @param configFilePath
	 *            配置文件路径
	 * @param excelFile
	 *            要检查的excel文件
	 * @param type
	 *            类型：城市公交，农村客运，出租汽车
	 * @throws Exception
	 */
	public void checkSysTemplate(Map<String, String> resultMap, String configFilePath, File excelFile, String type)
			throws Exception {
		logger.debug("---开始检查模板校验字符串---");
		Document document = ExcelImportUtil.getConfigFileDoc(configFilePath);
		ExcelImportUtil.genWorkbook(excelFile, document, type, resultMap);
		logger.debug("---结束检查模板校验字符串---");
	}

	/**
	 * 将update sql中的set之后的要更新的字段和值放到map中 update 语句形式为 UPDATE TABLENAME SET
	 * FIELD1=VALUE1,FIELD2=VALUE2...
	 * 
	 * @param updateSql
	 * @return
	 */
	public Map<String, String> updateToMap(String updateSql) {
		// logger.debug("---将update语句中的set部分放置到map中----");
		Map<String, String> resultMap = new HashMap<String, String>();
		String convertUpdate = updateSql.substring(updateSql.lastIndexOf("SET") + 3)
				.replace(ExcelConstants.SQL_TAIL, "").trim();
		// logger.debug("---去掉update语句头后---" + convertUpdate);
		// 去掉WHERE子句
		convertUpdate = convertUpdate.substring(0, convertUpdate.lastIndexOf("WHERE")).trim();
		// logger.debug("---去掉where子句后---" + convertUpdate);
		String[] fieldsValues = convertUpdate.split(",");
		for (String fv : fieldsValues) {
			String[] updatedFieldValue = fv.split("=");
			resultMap.put(updatedFieldValue[0], updatedFieldValue[1]);
		}
		return resultMap;
	}

	/**
	 * 将fieldValueMap中的字段名称和字段值添加到所有的insertSql中
	 * 
	 * @param insertSqls
	 *            生成的insert语句
	 * @param fieldValueMap
	 *            存放要添加到生成的insert语句中的field和value该map中key：字段名称,value:字段值
	 * @return
	 */
	public String[] remakeInsert(String insertSqls, Map<String, Object> fieldValueMap) throws Exception{
		return remakeInsert(insertSqls, fieldValueMap, null,null);
	}

	/**
	 * 将fieldValueMap中的字段名称和字段值添加到所有的insertSql中
	 * 
	 * @param insertSqls
	 *            生成的insert语句
	 * @param fieldValueMap
	 *            存放要添加到生成的insert语句中的field和value该map中key：字段名称,value:字段值
	 * @param pkName
	 *            主键名称
	 * @return
	 */
	public String[] remakeInsert(String insertSqls, Map<String, Object> fieldValueMap, String pkName, DBAccess dbAccess) throws Exception{
		String[] inserts = insertSqls.split(ExcelConstants.SQL_TAIL);
		Connection conn = dbAccess.getConnection();
		try {
			for (int i = 0; i < inserts.length; i++) {
				logger.debug("---分割后的insert语句---" + inserts[i]);
				//预先重组与业务相关
				inserts[i] = preRemake(inserts[i]);
				
				String firstPart = inserts[i].substring(0, inserts[i].lastIndexOf(ExcelConstants.SQL_INSERT_VALUE_FLAG));
				String secPart = inserts[i].substring(inserts[i].lastIndexOf(ExcelConstants.SQL_INSERT_VALUE_FLAG),
						inserts[i].lastIndexOf(")"));
				
				// 与业务相关的字段名称和值添加到insert中
				for (Map.Entry<String, Object> entry : fieldValueMap.entrySet()) {
					firstPart += "," + entry.getKey();
					if (entry.getValue() instanceof Long || entry.getValue() instanceof Float || entry.getValue() instanceof Double ) {
						secPart += "," + entry.getValue();
					} else if (entry.getValue() instanceof String) {
						secPart += "," +"'"+ entry.getValue()+"'";
					}
				}
				
				if (StringUtils.isNotBlank(pkName)) {
					firstPart += ","+pkName;
					secPart +=","+dbAccess.getSequence(conn);
				}
				// 添加完毕后重组insert语句
				inserts[i] = firstPart
						+ secPart.replace(ExcelConstants.SQL_INSERT_VALUE_FLAG, ExcelConstants.SQL_INSERT_VALUE) + ")";
				logger.debug("---重组后的insert---" + inserts[i]);
			}
		} catch (Exception e) {
			logger.error(e.getMessage(),e);
			throw e;
		}finally{
			dbAccess.release(conn);
		}
		return inserts;
	}
	
	/**
	 * 预先重组insert语句，该方法与业务相关，各模块需要重写
	 * @param insertSql 格式为：INSERT INTO TABLENAME (FIELD1,FIELD2...) @%&#VALUES#&%@ (VALUE1,VALUE2)
	 * @return 预先重组完成后的insert语句
	 */
	public String preRemake(String insertSql){
		String result = insertSql;
		return result;
	}
	/**
	 * 备份excel文件到指定目录
	 * 
	 * @param excelFile
	 *            要备份的excel文件
	 * @param backupDir
	 *            备份目录
	 * @param isSuccess
	 *            备份的文件是否为导入成功的文件：体现在最后重命名的文件名中
	 */
	public void backupFile(File excelFile, String backupDir, boolean isSuccess) {
		logger.debug("---移动---" + excelFile.getName() + "---到---" + backupDir);
		File bDir = new File(backupDir);
		try {
			FileUtils.moveFileToDirectory(excelFile, bDir, true);
			String backupFileName = backupDir + excelFile.getName();
			if (isSuccess) {// 解析并导入成功的文件
				backupFileName = backupFileName + "_success";
				logger.info("---解析文件成功---备份文件名称：" + backupFileName + "\n");
				FileUtils.moveFile(new File(backupDir + excelFile.getName()), new File(backupFileName));
			} else {// 解析过程中失败的文件
				backupFileName = backupFileName + "_fail";
				logger.info("---解析文件失败---备份文件名称：" + backupFileName + "\n");
				FileUtils.moveFile(new File(backupDir + excelFile.getName()), new File(backupFileName));
			}

		} catch (IOException e) {
			logger.error(e.getMessage(), e);
		}
	}

	/**
	 * 获得当前时间固定格式yyyyMMdd
	 * 
	 * @return 返回当前时间yyyyMMdd格式
	 */
	public String getNowDateString() {
		return getNowDateString("yyyyMMdd");
	}

	/**
	 * 指定日期类型返回日期字符串
	 * 
	 * @param datePattern
	 * @return
	 */
	public String getNowDateString(String datePattern) {
		SimpleDateFormat sdf = new SimpleDateFormat(datePattern);
		return sdf.format(new Date());
	}
	
	/**
	 * 获得insert语句中的values后面部分
	 * @param insertSql
	 * @return values后面部分
	 */
	public String getSqlValuePart(String insertSql) {
		String tempVPart = insertSql.substring(insertSql.lastIndexOf(ExcelConstants.SQL_INSERT_VALUE_FLAG)).replace(ExcelConstants.SQL_INSERT_VALUE_FLAG,"");
		String valuePart = tempVPart.substring(0, StringUtils.lastIndexOf(tempVPart, ")"));// insert的值部分
		return valuePart;
	}

	/**
	 * 获得insert语句中的前缀即 INSERT INTO TABLENAME 部分
	 * @param insertSql
	 * @return
	 */
	public String getSqlInsertPrefix(String insertSql) {
		String tempSPart = insertSql.substring(0, insertSql.lastIndexOf(ExcelConstants.SQL_INSERT_VALUE_FLAG));
		String insertPreFix = tempSPart.substring(0, StringUtils.indexOf(tempSPart, "(")+1);// INSERT INTO TABLENAME 部分
		return insertPreFix;
	}

	/**
	 * 获得insert语句中的声明部分即：insert语句中的tablename后面 values 前面的部分（不包括括号）
	 * @param insertSql
	 * @return
	 */
	public String getSqlStatePart(String insertSql) {
		String tempSPart = insertSql.substring(0, insertSql.lastIndexOf(ExcelConstants.SQL_INSERT_VALUE_FLAG));
		String statePart = tempSPart.substring(StringUtils.indexOf(tempSPart, "(")).replace("(", "");// insert声明部分
		return statePart;
	}

}
