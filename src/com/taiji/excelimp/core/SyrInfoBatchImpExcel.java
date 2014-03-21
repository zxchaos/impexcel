package com.taiji.excelimp.core;

import java.io.File;
import java.sql.Connection;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.dom4j.Document;

import com.taiji.excelimp.db.DBAccess;
import com.taiji.excelimp.util.ExcelConstants;
import com.taiji.excelimp.util.ExcelImportUtil;

public class SyrInfoBatchImpExcel extends AbstractImpExcel {
	private Workbook workbook;
	private String hylb;
	@Override
	public void importExcel(Properties sysConfig, DBAccess dbAccess) throws Exception {
		logger.debug("---受益人信息批量导入开始---");
		boolean isSuccess = false;
		String baseDir = sysConfig.getProperty("impDir");
		File registerDir = new File(baseDir + File.separator + sysConfig.getProperty("syrInfoBatchImpDirName"));
		String backupDir = sysConfig.getProperty("backupDir") + File.separator + getNowDateString() + File.separator;

		File[] impFiles = super.checkImpDir(registerDir.getAbsolutePath());
		if (impFiles == null) {
			return;
		}
		
		for (int i = 0; i < impFiles.length; i++) {
			String fileName = impFiles[i].getName();
			// 受益人信息导入的Excel文件命名规范：UUID_操作类别_单位Id.xlsx(.xls)
			// 操作类别包括：csgjsyr,nckysyr,czqcsyr
			// 也对应着配置文件中的template元素中的templateId属性
			String excelFileName = fileName.substring(0, fileName.lastIndexOf("."));
			String[] fileNameParts = excelFileName.split("_");
			String insertSqls = "";
			hylb = fileNameParts[1].substring(0, fileNameParts[1].lastIndexOf("syr"));
			String templateId = excelFileName.split("_")[1];
			String dwid = fileNameParts[2];
			Map<String, String> resultMap = new HashMap<String, String>();
			Map<String, String> infoFieldMap = new HashMap<String, String>();
			infoFieldMap.put("hylb", hylb);
			infoFieldMap.put("dwid", dwid);
			workbook = null;
			Document document = null;
			try {
				String configFilePath = sysConfig.getProperty("configFilePath");
				document = ExcelImportUtil.getConfigFileDoc(configFilePath);

				// 获得要导入的文件的工作簿对象并检查文件的有效性
				workbook = ExcelImportUtil.genWorkbook(impFiles[i], document, templateId, resultMap);

				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					// 开始解析文件
					ExcelImportUtil.importExcel(workbook, document, templateId, resultMap);
				}
				
				insertSqls = resultMap.get(ExcelConstants.SQLS_KEY);
				if (!ExcelConstants.FAIL.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY)) && StringUtils.isBlank(insertSqls)) {
					logger.info("+++生成的sql为空+++可能是模板中没有数据");
					ExcelImportUtil.setFailMsg(resultMap, "导入的模板中不包含数据");
				}
				
				String [] inserts = null;
				List<Map<String, String>> jcbList = new ArrayList<Map<String,String>>();
				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					inserts = insertSqls.split(ExcelConstants.SQL_TAIL);
					jcbList = getJcbList(resultMap,sysConfig,inserts.length);
				}
				
				
				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					// 若生成insert语句成功则执行插入操作
					Map<String, Object> fieldValueMap = new HashMap<String, Object>();
					fieldValueMap.put("QYID", dwid);
					doMultInsert(inserts, fieldValueMap, "ID", dbAccess,sysConfig,jcbList);
					super.insertImpInfo(dbAccess, resultMap, infoFieldMap, true, super.getType());
					isSuccess = true;
				} else {
					// 若生成失败将错误信息写入数据库
					isSuccess = false;
					super.insertImpInfo(dbAccess, resultMap, infoFieldMap, isSuccess, super.getType());
				}
			} catch (Exception e) {
				logger.info("+++导入出现异常+++");
				logger.error(e.getMessage(), e);
				isSuccess = false;
				resultMap.remove(ExcelConstants.MSG_KEY);
				ExcelImportUtil.setFailMsg(resultMap, "导入异常,请联系管理人员");
				super.insertImpInfo(dbAccess, resultMap, infoFieldMap, isSuccess, super.getType());
				logger.info("+++继续导入下一个文件+++");
				continue;
			} finally {
				// 将处理完成的文件移动到备份目录
				backupFile(impFiles[i], backupDir, isSuccess);
			}
		}
	}

	/**
	 * 获得模板中的每一行的前三列即：车牌号码，车牌颜色，变更情况的list
	 * @param resultMap
	 * @return
	 * @throws Exception
	 */
	private List<Map<String, String>> getJcbList(Map<String, String> resultMap, Properties config, Integer sqlCount) throws Exception {
		List<Map<String, String>> jcbList = new ArrayList<Map<String,String>>();
		logger.debug("---获取工作簿中的车牌号码-车牌颜色-变更情况---");
		if (workbook != null) {
			Sheet sheet = workbook.getSheet(hylb+"syr");
			Integer startNum = Integer.valueOf(config.getProperty("syrInfoDataRowStartNum"));
			for (int i = 0; i < sqlCount; i++) {
				Map<String, String>jcbMap = new HashMap<String, String>();
				Row row = sheet.getRow(i+startNum-1);
				logger.debug("---解析行---"+row.getRowNum());
				String cphmValue = ExcelImportUtil.getCellValue(row.getCell(0));
				String cpysValue = ExcelImportUtil.getCellValue(row.getCell(1));
				String bgqkValue = ExcelImportUtil.getCellValue(row.getCell(2));
				if (StringUtils.isBlank(cphmValue)) {
					String failMsg = "第"+(row.getRowNum()+1)+"行，第A列，不能为空";
					logger.debug(failMsg);
					ExcelImportUtil.setFailMsg(resultMap, failMsg);
					continue;
				}
				if (StringUtils.isBlank(cpysValue)) {
					String failMsg = "第"+(row.getRowNum()+1)+"行，第B列，不能为空";
					logger.debug(failMsg);
					ExcelImportUtil.setFailMsg(resultMap, failMsg);
					continue;
				}
				if (StringUtils.isBlank(bgqkValue)) {
					String failMsg = "第"+(row.getRowNum()+1)+"行，第C列，不能为空";
					logger.debug(failMsg);
					ExcelImportUtil.setFailMsg(resultMap, failMsg);
					continue;
				}
				
				jcbMap.put("cphm", cphmValue);//车牌号码
				jcbMap.put("cpys", cpysValue);//车牌颜色
				jcbMap.put("bgqk", bgqkValue);//变更情况
				jcbList.add(jcbMap);
				logger.debug("---工作簿---车牌号码--"+cphmValue+"---颜色--"+cpysValue+"---变更情况--"+bgqkValue+"---读取完毕---");
			}
		}else {
			throw new Exception("工作簿为空！");
		}
		return jcbList;
	}
	
	/**
	 * 重组insertSqls 并执行 重组后的insert语句并且向基础表中插入插入的受益人id
	 * 
	 * @param insertSqls
	 *            生成的insert语句
	 * @param fieldValueMap
	 *            存放要添加到生成的insert语句中的field和value该map中key：字段名称,value:字段值
	 * @param pkName
	 *            主键名称
	 * @param dbAccess 数据库访问对象
	 * @param config 系统配置对象
	 * @return
	 */
	public void doMultInsert(String [] inserts, Map<String, Object> fieldValueMap, String pkName, DBAccess dbAccess, Properties config, List<Map<String, String>>jcbList) throws Exception{
		
		Connection conn = dbAccess.getConnection();
		try {
			conn.setAutoCommit(false);
			for (int i = 0; i < inserts.length; i++) {
				logger.debug("---multinsert---获得单个insert语句---" + inserts[i]);
				String prefix = getMultInsertSqlPrefix(inserts[i]);
				String [] selects = getMultInsertsSelects(inserts[i], prefix);
				prefix = prefix.replace(" (", " (ID,QYID,");
				logger.debug("---multinsert---添加完主键id---qyid后---prefix---"+prefix);
				String selectPart = "";
				String pks="";
				
				for (int j = 0; j < selects.length; j++) {
					long pk = dbAccess.getSequence(conn);
					String noSelect = selects[j].replace("SELECT@%&", "");
					noSelect = pk+","+fieldValueMap.get("QYID")+","+noSelect;
					selects[j] = "SELECT@%&"+noSelect;
					if (logger.isDebugEnabled()) {
						logger.debug("---multinsert---生成主键---增加单位id---后的select为---"+selects[j]);
					}
					pks += pk+",";
					selectPart+= selects[j]+"@%&UNION@%&";
				}
				selectPart = selectPart.substring(0, selectPart.lastIndexOf("@%&UNION@%&"));
				pks = pks.substring(0, pks.lastIndexOf(","));
				inserts[i] = prefix+ " " + selectPart;
				logger.debug("---multinsert---重组后的insert语句为"+inserts[i]);
				
				Map<String, String> jcbMap = jcbList.get(i);
				String filter = " CPHM='"+jcbMap.get("cphm")+"' AND CPYS='"+jcbMap.get("cpys")+"' AND BGQK='"+jcbMap.get("bgqk")+"'";
				String order = " SJID desc";
				String sql = "SELECT SJID FROM "+hylb+"JCB WHERE"+filter+" ORDER BY "+order;
				logger.debug("---multinsert---查询基础表的sql---"+sql);
				String sjid = dbAccess.getOneFieldContent(sql, conn);
				inserts[i] = inserts[i].replace(ExcelConstants.SQL_MULTINSERT_FROM_DUAL_FLAG, ExcelConstants.SQL_MULTINSERT_FROM_DUAL).replace(ExcelConstants.SQL_MULTINSERT_UNION_SELECT_FLAG, ExcelConstants.SQL_MULTINSERT_UNION_SELECT).replace("SELECT@%&", "SELECT ");
				logger.debug("---multinsert---替换完成标识占位后的sql---"+inserts[i]);
				dbAccess.multInsertSyrAndJCB(conn, inserts[i], pks, sjid, hylb);
			}
			conn.commit();
		} catch (Exception e) {
			logger.error(e.getMessage(),e);
			conn.rollback();
			throw e;
		}finally{
			dbAccess.release(conn);
		}
	}

	/**
	 * 获得一条insert多条记录的sql前缀
	 * @param multInsertSql
	 * @return
	 */
	public String getMultInsertSqlPrefix(String multInsertSql){
		String result = "";
		result = multInsertSql.substring(0, multInsertSql.indexOf(") SELECT@%&")+1);
		logger.debug("---multinsert---prefix---"+result);
		return result;
	}
	
	/**
	 * 获得一条insert插入多条记录语句中的select子句
	 * @param multInsertSql
	 * @param prefix
	 * @return
	 */
	public String [] getMultInsertsSelects(String multInsertSql, String prefix){
		String selectPart = multInsertSql.replace(prefix, "");
		String [] result = selectPart.split("@%&UNION@%&");
		if (logger.isDebugEnabled()) {
			logger.debug("---multinsert---select部分分割后---");
			for (int i = 0; i < result.length; i++) {
				logger.debug(result[i]);
			}
		}
		return result;
	}
}