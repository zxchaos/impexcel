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
	/**ģ�����Ŀǰtypeȡ�ĸ�ֵ��1�������������룬2������ע����Ϣ�������룬3��ũ�������·�������룬4����������Ϣ��������*/
	private String type;

	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
	}

	/**
	 * ����excelDirĿ¼�µ�excel�ļ�
	 * 
	 * @param excelDir
	 */
	public abstract void importExcel(Properties sysConfig, DBAccess dbAccess) throws Exception;

	/**
	 * @param impDir
	 * @return impDir�µ��ļ� ����Ŀ¼�����ļ��򷵻�null
	 * @throws Exception
	 */
	public File[] checkImpDir(String impDir) throws Exception {
		File excelDir = new File(impDir);
		if (!excelDir.exists()) {
			// ����Ŀ¼���������Զ�����
			excelDir.mkdirs();
		} else if (!excelDir.isDirectory()) {
			logger.error("·����" + impDir + "����Ŀ¼");
			throw new Exception("·����" + impDir + "����Ŀ¼");
		}
		logger.info("+++��ѯĿ¼+++" + excelDir);
		// ѡ��excel 2003 �� 2007���ļ�
		File[] excelFiles = ExcelImportUtil.getDirFiles(excelDir);
		if (excelFiles.length == 0) {
			logger.info("+++Ŀ¼" + excelDir.getAbsolutePath() + "�����ļ�+++");
			return null;
		}
		logger.info("+++Ŀ¼�е�Excel�ļ�����+++" + excelFiles.length);
		return excelFiles;
	}

	/**
	 * ��������Ϣд����־
	 * 
	 * @param dbAccess
	 * @param resultMap
	 * @param infoFieldsMap Ҫ������־���ֶ��Լ�ֵ��map��key��Ϊ�ֶ����ƣ�value���ֶ�ֵ
	 * @param type ģ�����ͣ�1����������¼�룻2������ע����Ϣ��������
	 *            
	 */
	public void insertImpInfo(DBAccess dbAccess, Map<String, String> resultMap, Map<String, String> infoFieldsMap,
			boolean isSuccess,String type) throws Exception {
		logger.debug("---����������Ϣд�����ݿ�---");
		long dwid = Long.valueOf(infoFieldsMap.get("dwid"));
		String hylb = infoFieldsMap.get("hylb");
		String sj = getNowDateString("yyyy��MM��dd��HHʱmm��");
		String info = "";
		if (isSuccess) {
			info = "���������Ϣ�ɹ�";
		} else {
			info = resultMap.get(ExcelConstants.MSG_KEY);
		}
		logger.info("+++��������Ϣ+++\n" + info);
		if (StringUtils.isNotBlank(info) && info.length()>2000) {//��������Ϣ���ȳ���2000���Զ��ض�Ϊ2000����
			logger.info("+++������Ϣ����2000���ַ�+++");
			info = info.substring(0,2001);
			logger.debug("---�ضϳ������ȵĴ�����Ϣ��---"+info);
		}
//		dbAccess.insertErrorInfo(dwid, hylb, sj, info, type);
		logger.debug("---��������Ϣд�����ݿ����---");
	}

	/**
	 * ����ļ��Ƿ�Ϊϵͳ�ṩ��ģ��
	 * 
	 * @param resultMap
	 *            �����
	 * @param configFilePath
	 *            �����ļ�·��
	 * @param excelFile
	 *            Ҫ����excel�ļ�
	 * @param type
	 *            ���ͣ����й�����ũ����ˣ���������
	 * @throws Exception
	 */
	public void checkSysTemplate(Map<String, String> resultMap, String configFilePath, File excelFile, String type)
			throws Exception {
		logger.debug("---��ʼ���ģ��У���ַ���---");
		Document document = ExcelImportUtil.getConfigFileDoc(configFilePath);
		ExcelImportUtil.genWorkbook(excelFile, document, type, resultMap);
		logger.debug("---�������ģ��У���ַ���---");
	}

	/**
	 * ��update sql�е�set֮���Ҫ���µ��ֶκ�ֵ�ŵ�map�� update �����ʽΪ UPDATE TABLENAME SET
	 * FIELD1=VALUE1,FIELD2=VALUE2...
	 * 
	 * @param updateSql
	 * @return
	 */
	public Map<String, String> updateToMap(String updateSql) {
		// logger.debug("---��update����е�set���ַ��õ�map��----");
		Map<String, String> resultMap = new HashMap<String, String>();
		String convertUpdate = updateSql.substring(updateSql.lastIndexOf("SET") + 3)
				.replace(ExcelConstants.SQL_TAIL, "").trim();
		// logger.debug("---ȥ��update���ͷ��---" + convertUpdate);
		// ȥ��WHERE�Ӿ�
		convertUpdate = convertUpdate.substring(0, convertUpdate.lastIndexOf("WHERE")).trim();
		// logger.debug("---ȥ��where�Ӿ��---" + convertUpdate);
		String[] fieldsValues = convertUpdate.split(",");
		for (String fv : fieldsValues) {
			String[] updatedFieldValue = fv.split("=");
			resultMap.put(updatedFieldValue[0], updatedFieldValue[1]);
		}
		return resultMap;
	}

	/**
	 * ��fieldValueMap�е��ֶ����ƺ��ֶ�ֵ��ӵ����е�insertSql��
	 * 
	 * @param insertSqls
	 *            ���ɵ�insert���
	 * @param fieldValueMap
	 *            ���Ҫ��ӵ����ɵ�insert����е�field��value��map��key���ֶ�����,value:�ֶ�ֵ
	 * @return
	 */
	public String[] remakeInsert(String insertSqls, Map<String, Object> fieldValueMap) throws Exception{
		return remakeInsert(insertSqls, fieldValueMap, null,null);
	}

	/**
	 * ��fieldValueMap�е��ֶ����ƺ��ֶ�ֵ��ӵ����е�insertSql��
	 * 
	 * @param insertSqls
	 *            ���ɵ�insert���
	 * @param fieldValueMap
	 *            ���Ҫ��ӵ����ɵ�insert����е�field��value��map��key���ֶ�����,value:�ֶ�ֵ
	 * @param pkName
	 *            ��������
	 * @return
	 */
	public String[] remakeInsert(String insertSqls, Map<String, Object> fieldValueMap, String pkName, DBAccess dbAccess) throws Exception{
		String[] inserts = insertSqls.split(ExcelConstants.SQL_TAIL);
		Connection conn = dbAccess.getConnection();
		try {
			for (int i = 0; i < inserts.length; i++) {
				logger.debug("---�ָ���insert���---" + inserts[i]);
				//Ԥ��������ҵ�����
				inserts[i] = preRemake(inserts[i]);
				
				String firstPart = inserts[i].substring(0, inserts[i].lastIndexOf(ExcelConstants.SQL_INSERT_VALUE_FLAG));
				String secPart = inserts[i].substring(inserts[i].lastIndexOf(ExcelConstants.SQL_INSERT_VALUE_FLAG),
						inserts[i].lastIndexOf(")"));
				
				// ��ҵ����ص��ֶ����ƺ�ֵ��ӵ�insert��
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
				// �����Ϻ�����insert���
				inserts[i] = firstPart
						+ secPart.replace(ExcelConstants.SQL_INSERT_VALUE_FLAG, ExcelConstants.SQL_INSERT_VALUE) + ")";
				logger.debug("---������insert---" + inserts[i]);
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
	 * Ԥ������insert��䣬�÷�����ҵ����أ���ģ����Ҫ��д
	 * @param insertSql ��ʽΪ��INSERT INTO TABLENAME (FIELD1,FIELD2...) @%&#VALUES#&%@ (VALUE1,VALUE2)
	 * @return Ԥ��������ɺ��insert���
	 */
	public String preRemake(String insertSql){
		String result = insertSql;
		return result;
	}
	/**
	 * ����excel�ļ���ָ��Ŀ¼
	 * 
	 * @param excelFile
	 *            Ҫ���ݵ�excel�ļ�
	 * @param backupDir
	 *            ����Ŀ¼
	 * @param isSuccess
	 *            ���ݵ��ļ��Ƿ�Ϊ����ɹ����ļ���������������������ļ�����
	 */
	public void backupFile(File excelFile, String backupDir, boolean isSuccess) {
		logger.debug("---�ƶ�---" + excelFile.getName() + "---��---" + backupDir);
		File bDir = new File(backupDir);
		try {
			FileUtils.moveFileToDirectory(excelFile, bDir, true);
			String backupFileName = backupDir + excelFile.getName();
			if (isSuccess) {// ����������ɹ����ļ�
				backupFileName = backupFileName + "_success";
				logger.info("---�����ļ��ɹ�---�����ļ����ƣ�" + backupFileName + "\n");
				FileUtils.moveFile(new File(backupDir + excelFile.getName()), new File(backupFileName));
			} else {// ����������ʧ�ܵ��ļ�
				backupFileName = backupFileName + "_fail";
				logger.info("---�����ļ�ʧ��---�����ļ����ƣ�" + backupFileName + "\n");
				FileUtils.moveFile(new File(backupDir + excelFile.getName()), new File(backupFileName));
			}

		} catch (IOException e) {
			logger.error(e.getMessage(), e);
		}
	}

	/**
	 * ��õ�ǰʱ��̶���ʽyyyyMMdd
	 * 
	 * @return ���ص�ǰʱ��yyyyMMdd��ʽ
	 */
	public String getNowDateString() {
		return getNowDateString("yyyyMMdd");
	}

	/**
	 * ָ���������ͷ��������ַ���
	 * 
	 * @param datePattern
	 * @return
	 */
	public String getNowDateString(String datePattern) {
		SimpleDateFormat sdf = new SimpleDateFormat(datePattern);
		return sdf.format(new Date());
	}
	
	/**
	 * ���insert����е�values���沿��
	 * @param insertSql
	 * @return values���沿��
	 */
	public String getSqlValuePart(String insertSql) {
		String tempVPart = insertSql.substring(insertSql.lastIndexOf(ExcelConstants.SQL_INSERT_VALUE_FLAG)).replace(ExcelConstants.SQL_INSERT_VALUE_FLAG,"");
		String valuePart = tempVPart.substring(0, StringUtils.lastIndexOf(tempVPart, ")"));// insert��ֵ����
		return valuePart;
	}

	/**
	 * ���insert����е�ǰ׺�� INSERT INTO TABLENAME ����
	 * @param insertSql
	 * @return
	 */
	public String getSqlInsertPrefix(String insertSql) {
		String tempSPart = insertSql.substring(0, insertSql.lastIndexOf(ExcelConstants.SQL_INSERT_VALUE_FLAG));
		String insertPreFix = tempSPart.substring(0, StringUtils.indexOf(tempSPart, "(")+1);// INSERT INTO TABLENAME ����
		return insertPreFix;
	}

	/**
	 * ���insert����е��������ּ���insert����е�tablename���� values ǰ��Ĳ��֣����������ţ�
	 * @param insertSql
	 * @return
	 */
	public String getSqlStatePart(String insertSql) {
		String tempSPart = insertSql.substring(0, insertSql.lastIndexOf(ExcelConstants.SQL_INSERT_VALUE_FLAG));
		String statePart = tempSPart.substring(StringUtils.indexOf(tempSPart, "(")).replace("(", "");// insert��������
		return statePart;
	}

}
