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
		logger.debug("---��������Ϣ�������뿪ʼ---");
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
			// ��������Ϣ�����Excel�ļ������淶��UUID_�������_��λId.xlsx(.xls)
			// ������������csgjsyr,nckysyr,czqcsyr
			// Ҳ��Ӧ�������ļ��е�templateԪ���е�templateId����
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

				// ���Ҫ������ļ��Ĺ��������󲢼���ļ�����Ч��
				workbook = ExcelImportUtil.genWorkbook(impFiles[i], document, templateId, resultMap);

				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					// ��ʼ�����ļ�
					ExcelImportUtil.importExcel(workbook, document, templateId, resultMap);
				}
				
				insertSqls = resultMap.get(ExcelConstants.SQLS_KEY);
				if (!ExcelConstants.FAIL.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY)) && StringUtils.isBlank(insertSqls)) {
					logger.info("+++���ɵ�sqlΪ��+++������ģ����û������");
					ExcelImportUtil.setFailMsg(resultMap, "�����ģ���в���������");
				}
				
				String [] inserts = null;
				List<Map<String, String>> jcbList = new ArrayList<Map<String,String>>();
				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					inserts = insertSqls.split(ExcelConstants.SQL_TAIL);
					jcbList = getJcbList(resultMap,sysConfig,inserts.length);
				}
				
				
				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					// ������insert���ɹ���ִ�в������
					Map<String, Object> fieldValueMap = new HashMap<String, Object>();
					fieldValueMap.put("QYID", dwid);
					doMultInsert(inserts, fieldValueMap, "ID", dbAccess,sysConfig,jcbList);
					super.insertImpInfo(dbAccess, resultMap, infoFieldMap, true, super.getType());
					isSuccess = true;
				} else {
					// ������ʧ�ܽ�������Ϣд�����ݿ�
					isSuccess = false;
					super.insertImpInfo(dbAccess, resultMap, infoFieldMap, isSuccess, super.getType());
				}
			} catch (Exception e) {
				logger.info("+++��������쳣+++");
				logger.error(e.getMessage(), e);
				isSuccess = false;
				resultMap.remove(ExcelConstants.MSG_KEY);
				ExcelImportUtil.setFailMsg(resultMap, "�����쳣,����ϵ������Ա");
				super.insertImpInfo(dbAccess, resultMap, infoFieldMap, isSuccess, super.getType());
				logger.info("+++����������һ���ļ�+++");
				continue;
			} finally {
				// ��������ɵ��ļ��ƶ�������Ŀ¼
				backupFile(impFiles[i], backupDir, isSuccess);
			}
		}
	}

	/**
	 * ���ģ���е�ÿһ�е�ǰ���м������ƺ��룬������ɫ����������list
	 * @param resultMap
	 * @return
	 * @throws Exception
	 */
	private List<Map<String, String>> getJcbList(Map<String, String> resultMap, Properties config, Integer sqlCount) throws Exception {
		List<Map<String, String>> jcbList = new ArrayList<Map<String,String>>();
		logger.debug("---��ȡ�������еĳ��ƺ���-������ɫ-������---");
		if (workbook != null) {
			Sheet sheet = workbook.getSheet(hylb+"syr");
			Integer startNum = Integer.valueOf(config.getProperty("syrInfoDataRowStartNum"));
			for (int i = 0; i < sqlCount; i++) {
				Map<String, String>jcbMap = new HashMap<String, String>();
				Row row = sheet.getRow(i+startNum-1);
				logger.debug("---������---"+row.getRowNum());
				String cphmValue = ExcelImportUtil.getCellValue(row.getCell(0));
				String cpysValue = ExcelImportUtil.getCellValue(row.getCell(1));
				String bgqkValue = ExcelImportUtil.getCellValue(row.getCell(2));
				if (StringUtils.isBlank(cphmValue)) {
					String failMsg = "��"+(row.getRowNum()+1)+"�У���A�У�����Ϊ��";
					logger.debug(failMsg);
					ExcelImportUtil.setFailMsg(resultMap, failMsg);
					continue;
				}
				if (StringUtils.isBlank(cpysValue)) {
					String failMsg = "��"+(row.getRowNum()+1)+"�У���B�У�����Ϊ��";
					logger.debug(failMsg);
					ExcelImportUtil.setFailMsg(resultMap, failMsg);
					continue;
				}
				if (StringUtils.isBlank(bgqkValue)) {
					String failMsg = "��"+(row.getRowNum()+1)+"�У���C�У�����Ϊ��";
					logger.debug(failMsg);
					ExcelImportUtil.setFailMsg(resultMap, failMsg);
					continue;
				}
				
				jcbMap.put("cphm", cphmValue);//���ƺ���
				jcbMap.put("cpys", cpysValue);//������ɫ
				jcbMap.put("bgqk", bgqkValue);//������
				jcbList.add(jcbMap);
				logger.debug("---������---���ƺ���--"+cphmValue+"---��ɫ--"+cpysValue+"---������--"+bgqkValue+"---��ȡ���---");
			}
		}else {
			throw new Exception("������Ϊ�գ�");
		}
		return jcbList;
	}
	
	/**
	 * ����insertSqls ��ִ�� ������insert��䲢����������в�������������id
	 * 
	 * @param insertSqls
	 *            ���ɵ�insert���
	 * @param fieldValueMap
	 *            ���Ҫ��ӵ����ɵ�insert����е�field��value��map��key���ֶ�����,value:�ֶ�ֵ
	 * @param pkName
	 *            ��������
	 * @param dbAccess ���ݿ���ʶ���
	 * @param config ϵͳ���ö���
	 * @return
	 */
	public void doMultInsert(String [] inserts, Map<String, Object> fieldValueMap, String pkName, DBAccess dbAccess, Properties config, List<Map<String, String>>jcbList) throws Exception{
		
		Connection conn = dbAccess.getConnection();
		try {
			conn.setAutoCommit(false);
			for (int i = 0; i < inserts.length; i++) {
				logger.debug("---multinsert---��õ���insert���---" + inserts[i]);
				String prefix = getMultInsertSqlPrefix(inserts[i]);
				String [] selects = getMultInsertsSelects(inserts[i], prefix);
				prefix = prefix.replace(" (", " (ID,QYID,");
				logger.debug("---multinsert---���������id---qyid��---prefix---"+prefix);
				String selectPart = "";
				String pks="";
				
				for (int j = 0; j < selects.length; j++) {
					long pk = dbAccess.getSequence(conn);
					String noSelect = selects[j].replace("SELECT@%&", "");
					noSelect = pk+","+fieldValueMap.get("QYID")+","+noSelect;
					selects[j] = "SELECT@%&"+noSelect;
					if (logger.isDebugEnabled()) {
						logger.debug("---multinsert---��������---���ӵ�λid---���selectΪ---"+selects[j]);
					}
					pks += pk+",";
					selectPart+= selects[j]+"@%&UNION@%&";
				}
				selectPart = selectPart.substring(0, selectPart.lastIndexOf("@%&UNION@%&"));
				pks = pks.substring(0, pks.lastIndexOf(","));
				inserts[i] = prefix+ " " + selectPart;
				logger.debug("---multinsert---������insert���Ϊ"+inserts[i]);
				
				Map<String, String> jcbMap = jcbList.get(i);
				String filter = " CPHM='"+jcbMap.get("cphm")+"' AND CPYS='"+jcbMap.get("cpys")+"' AND BGQK='"+jcbMap.get("bgqk")+"'";
				String order = " SJID desc";
				String sql = "SELECT SJID FROM "+hylb+"JCB WHERE"+filter+" ORDER BY "+order;
				logger.debug("---multinsert---��ѯ�������sql---"+sql);
				String sjid = dbAccess.getOneFieldContent(sql, conn);
				inserts[i] = inserts[i].replace(ExcelConstants.SQL_MULTINSERT_FROM_DUAL_FLAG, ExcelConstants.SQL_MULTINSERT_FROM_DUAL).replace(ExcelConstants.SQL_MULTINSERT_UNION_SELECT_FLAG, ExcelConstants.SQL_MULTINSERT_UNION_SELECT).replace("SELECT@%&", "SELECT ");
				logger.debug("---multinsert---�滻��ɱ�ʶռλ���sql---"+inserts[i]);
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
	 * ���һ��insert������¼��sqlǰ׺
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
	 * ���һ��insert���������¼����е�select�Ӿ�
	 * @param multInsertSql
	 * @param prefix
	 * @return
	 */
	public String [] getMultInsertsSelects(String multInsertSql, String prefix){
		String selectPart = multInsertSql.replace(prefix, "");
		String [] result = selectPart.split("@%&UNION@%&");
		if (logger.isDebugEnabled()) {
			logger.debug("---multinsert---select���ַָ��---");
			for (int i = 0; i < result.length; i++) {
				logger.debug(result[i]);
			}
		}
		return result;
	}
}
