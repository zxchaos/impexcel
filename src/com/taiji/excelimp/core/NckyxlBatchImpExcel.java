package com.taiji.excelimp.core;

import java.io.File;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.dom4j.Document;

import com.taiji.excelimp.db.DBAccess;
import com.taiji.excelimp.util.ExcelConstants;
import com.taiji.excelimp.util.ExcelImportUtil;
/**
 * ũ�������·��������ҵ����
 * @author zhangxin
 *
 */
public class NckyxlBatchImpExcel extends AbstractImpExcel {
	@Override
	public void importExcel(Properties sysConfig, DBAccess dbAccess) throws Exception {
		logger.debug("---ũ�������·�������뿪ʼ---");
		boolean isSuccess = false;
		String baseDir = sysConfig.getProperty("impDir");
		File nckyDir = new File(baseDir + File.separator + sysConfig.getProperty("nckyxlBatchImpDirName"));
		String backupDir = sysConfig.getProperty("backupDir")+ File.separator + getNowDateString() + File.separator;
		File[] impFiles = checkImpDir(nckyDir.getAbsolutePath());
		if (impFiles == null) {
			return ;
		}
		
		for (int i = 0; i < impFiles.length; i++) {
			String fileName = impFiles[i].getName();
			// ������Ϣ�����Excel�ļ������淶��UUID_�������_��λId_���д���.xlsx(.xls)
			// ������������nckyxlgl
			// Ҳ��Ӧ�������ļ��е�templateԪ���е�templateId����
			String excelFileName = fileName.substring(0, fileName.lastIndexOf("."));
			String [] fileNameParts = excelFileName.split("_");
			String insertSqls = "";
			String hylb = fileNameParts[1].substring(0,fileNameParts[1].lastIndexOf("xlgl"));
			String templateId = excelFileName.split("_")[1];
			String dwid = fileNameParts[2];
			Map<String, String> resultMap = new HashMap<String, String>();
			Map<String, String>infoFieldMap = new HashMap<String, String>();
			infoFieldMap.put("hylb", hylb);
			infoFieldMap.put("dwid", dwid);
			Workbook workbook = null;
			Document document = null;
			try {
				String configFilePath = sysConfig.getProperty("configFilePath");
				document = ExcelImportUtil.getConfigFileDoc(configFilePath);
				
				//���Ҫ������ļ��Ĺ��������󲢼���ļ�����Ч��
				workbook = ExcelImportUtil.genWorkbook(impFiles[i], document, templateId, resultMap);
				
				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					//��ʼ�����ļ�
					ExcelImportUtil.importExcel(workbook, document, templateId, resultMap);
				}
				
				insertSqls = resultMap.get(ExcelConstants.SQLS_KEY);
				if (!ExcelConstants.FAIL.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY)) && StringUtils.isBlank(insertSqls)) {
					logger.info("+++���ɵ�sqlΪ��+++������ģ����û������");
					ExcelImportUtil.setFailMsg(resultMap, "�����ģ���в���������");
				}
				
				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					//������insert���ɹ���ִ�в������
					Map<String, Object>fieldValueMap = new HashMap<String, Object>();
					fieldValueMap.put("XS", Double.valueOf(0.1));
					fieldValueMap.put("ZT", Long.valueOf(0));
					fieldValueMap.put("SHENG", Long.valueOf(650000));
					fieldValueMap.put("SHI", Long.valueOf(fileNameParts[3]));//���ļ�������ȡ�õ��д���
					String [] inserts = super.remakeInsert(insertSqls, fieldValueMap,"BXID",dbAccess);
					//������������
					dbAccess.batchExecuteSqls(inserts);
					super.insertImpInfo(dbAccess, resultMap, infoFieldMap,true,super.getType());
					isSuccess = true;
				}else{
					//������ʧ�ܽ�������Ϣд�����ݿ�
					isSuccess = false;
					super.insertImpInfo(dbAccess, resultMap, infoFieldMap, isSuccess,super.getType());
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
			}finally{
				//��������ɵ��ļ��ƶ�������Ŀ¼
				backupFile(impFiles[i], backupDir, isSuccess);
			}
		}
	}

}
