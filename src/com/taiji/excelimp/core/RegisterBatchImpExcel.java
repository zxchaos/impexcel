package com.taiji.excelimp.core;

import java.io.File;
import java.math.BigDecimal;
import java.sql.Connection;
import java.text.ParseException;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.dom4j.Document;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.taiji.excelimp.db.DBAccess;
import com.taiji.excelimp.util.ExcelConstants;
import com.taiji.excelimp.util.ExcelImportUtil;

/**
 * ������Ϣ����ע�ᵼ�빦��
 *
 * @author zhangxin
 *
 */
public class RegisterBatchImpExcel extends AbstractImpExcel {

	public static Logger logger = LoggerFactory.getLogger(RegisterBatchImpExcel.class);

	@Override
	public void importExcel(Properties sysConfig, DBAccess dbAccess) throws Exception {
		logger.debug("---������Ϣ�������뿪ʼ---");
		boolean isSuccess = false;
		String baseDir = sysConfig.getProperty("impDir");
		File registerDir = new File(baseDir + File.separator + sysConfig.getProperty("registerBatchImpDirName"));
		String backupDir = sysConfig.getProperty("backupDir") + File.separator + getNowDateString() + File.separator;

		File[] impFiles = super.checkImpDir(registerDir.getAbsolutePath());
		if (impFiles == null) {
			return;
		}
		for (int i = 0; i < impFiles.length; i++) {
			String fileName = impFiles[i].getName();
			// ������Ϣ�����Excel�ļ������淶��UUID_�������_��λId_ʡ����_�д���_�ش���_��λ����.xlsx(.xls)
			// ������������csgjplzc,nckyplzc,czqcplzc
			// Ҳ��Ӧ�������ļ��е�templateԪ���е�templateId����
			String excelFileName = fileName.substring(0, fileName.lastIndexOf("."));
			String[] fileNameParts = excelFileName.split("_");
			String insertSqls = "";
			String hylb = fileNameParts[1].substring(0, fileNameParts[1].lastIndexOf("plzc"));
			String templateId = excelFileName.split("_")[1];
			String dwid = fileNameParts[2];
			Map<String, String> resultMap = new HashMap<String, String>();
			Map<String, String> infoFieldMap = new HashMap<String, String>();
			infoFieldMap.put("hylb", hylb);
			infoFieldMap.put("dwid", dwid);
			Workbook workbook = null;
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

				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					// ���ݳ��ƺ������
					isSuccess = !checkDuplicate(insertSqls, resultMap, dbAccess, hylb);
				}

				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY)) && ("csgj".equalsIgnoreCase(hylb)||"ncky".equalsIgnoreCase(hylb))) {
					// ������·���Ʋ�ѯ��·���е���·�Ƿ����
					insertSqls = xlmcCheck(insertSqls, resultMap, dbAccess, hylb);
					logger.debug("---��·�����ɺ�����sql---"+insertSqls);
				}

				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					// ������insert���ɹ���ִ�в������
					Map<String, Object> fieldValueMap = new HashMap<String, Object>();
					fieldValueMap.put("RZZT", Long.valueOf(3));
					fieldValueMap.put("CLDJLX", Long.valueOf(1));
					fieldValueMap.put("DWID", Long.valueOf(dwid));
					fieldValueMap.put("SHSHENG", Long.valueOf(fileNameParts[3]));
					fieldValueMap.put("SHSHI", Long.valueOf(fileNameParts[4]));
					fieldValueMap.put("SHXIAN", Long.valueOf(fileNameParts[5]));
					fieldValueMap.put("DWMC", fileNameParts[6]);
					String[] inserts = super.remakeInsert(insertSqls, fieldValueMap, "SJID", dbAccess);
					if (logger.isDebugEnabled()) {
						logger.debug("---�������ɵ�ִ��sql---");
						for (int j = 0; j < inserts.length; j++) {
							logger.debug(inserts[j]);
						}
					}
					// ������������
					dbAccess.batchExecuteSqls(inserts);
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
	 * �����·����
	 * @param insertSqls
	 * @param resultMap
	 * @param dbAccess
	 * @param hylb
	 * @return
	 */
	private String xlmcCheck(String insertSqls, Map<String, String>resultMap, DBAccess dbAccess, String hylb) throws Exception{
		StringBuffer result = new StringBuffer("");
		String[] inserts = insertSqls.split(ExcelConstants.SQL_TAIL);
		Connection conn = dbAccess.getConnection();
		try {
			for (int i = 0; i < inserts.length; i++) {
				String statePart = super.getSqlStatePart(inserts[i]);
				String valuePart = super.getSqlValuePart(inserts[i]);
				String insertPrefix = super.getSqlInsertPrefix(inserts[i]);
				String[] valueParts = valuePart.split(",");
				String xlmc = "";
				String tableName = "";
				if ("csgj".equalsIgnoreCase(hylb)) {
					xlmc = valueParts[13];// ��ó��й�����·����
					tableName = "T_CSGJ_XLGL";
				} else if ("ncky".equalsIgnoreCase(hylb)) {
					xlmc = valueParts[13];// ���ũ�������·����
					tableName = "T_NCKY_XLGL";
				}
				logger.debug("---��insert����л�õ���·����---"+xlmc);

				String bxId = getBXIDByXlmc(xlmc, conn, dbAccess, tableName);
				if (StringUtils.isNotBlank(bxId)) {// ����·���еļ�¼��������·id���뵽insert�����
					statePart += "," + "YYXLH";
					valuePart += "," + bxId;
					inserts[i] = insertPrefix + statePart + ExcelConstants.SQL_INSERT_VALUE_FLAG + valuePart + ")";
					logger.debug("---��·������---��·���д��ڼ�¼��" + bxId + "---�����insert���---" + inserts[i]);
					result.append(inserts[i]);
					result.append(ExcelConstants.SQL_TAIL);
				} else {
					String failMsg = "����Ϊ��" + xlmc + "����·������";
					logger.info(failMsg);
					ExcelImportUtil.setFailMsg(resultMap, failMsg);
					break;
				}

			}
		} catch (Exception e) {
			logger.error(e.getMessage(),e);
			throw e;
		} finally {
			dbAccess.release(conn);
		}
		return result.toString();
	}

	/**
	 * ������·���ƻ����·id
	 * @param xlmc
	 * @return
	 */
	private String getBXIDByXlmc(String xlmc, Connection conn, DBAccess dbAccess, String tableName) throws Exception{
		String result = "";
		String sql = "select BXID from "+tableName+" where XLMC ="+xlmc+"";
		result = dbAccess.getOneFieldContent(sql, conn);
		return result;
	}
	/**
	 * ���ݳ��ƺ��� ������ɫ ������ ��ѯ������Ϣ�Ƿ����ظ�
	 *
	 * @param insertSqls
	 *            ����excel���ɵ�insert���
	 * @param resultMap
	 *            �������
	 * @param dbAccess
	 *            ���ݿ���ʶ���
	 */
	private boolean checkDuplicate(String insertSqls, Map<String, String> resultMap, DBAccess dbAccess, String type) throws Exception{
		logger.debug("---��ѯ�ظ����ƺ���---");
		boolean result = false;
		String[] inserts = insertSqls.split(ExcelConstants.SQL_TAIL);
		Connection conn = dbAccess.getConnection();
		String tableName = "";
		if ("csgj".equalsIgnoreCase(type)) {
			tableName = "CSGJJCB";
		}else if ("ncky".equalsIgnoreCase(type)) {
			tableName = "NCKYJCB";
		}else if ("czqc".equalsIgnoreCase(type)) {
			tableName = "CZQCJCB";
		}else {
			throw new Exception("�޷�ȷ����ѯ�ظ����Ƶı�����");
		}
		try {
			for (int i = 0; i < inserts.length; i++) {
				logger.debug("---insert���---" + inserts[i]);
				String valuePart = inserts[i].substring(inserts[i].lastIndexOf(ExcelConstants.SQL_INSERT_VALUE_FLAG))
						.replace(ExcelConstants.SQL_INSERT_VALUE_FLAG, "");
				String [] valueParts = valuePart.split(",");
				// ���cphm�ֶζ�Ӧ��ֵ
				String cphmVal = valueParts[0].replace("'", "");
				//������ɫ
				String cpysVal = valueParts[1].replace("'", "");
				//������
				String bgqkVal = valueParts[2].replace("'", "");
				Map<String, String> valueMap = new HashMap<String, String>();
				valueMap.put("CPHM", cphmVal);
				valueMap.put("CPYS", cpysVal);
				valueMap.put("BGQK", bgqkVal);
				logger.debug("---���ƺ���---" + cphmVal+"---������ɫ---"+cpysVal+"---������---"+bgqkVal);
				if (dbAccess.isFieldValueDup(tableName, valueMap, conn)) {
					ExcelImportUtil.setFailMsg(resultMap, "���ƺ���Ϊ" + cphmVal + "��������ɫΪ"+cpysVal+"��������Ϊ"+bgqkVal+"�ļ�¼�Ѿ������޷�����");
					result = true;
				}
			}
		} catch (Exception e) {
			logger.error(e.getMessage(),e);
			throw e;
		}finally{
			dbAccess.release(conn);
		}
		return result;
	}

	@Override
	public String preRemake(String insertSql) {
		String result = "";
//		String[] insertParts = StringUtils.split(insertSql, ExcelConstants.SQL_INSERT_VALUE_FLAG);
		String statePart = super.getSqlStatePart(insertSql);
		String valuePart = super.getSqlValuePart(insertSql);
		String insertPreFix = super.getSqlInsertPrefix(insertSql);
		logger.debug("---statePart--"+statePart);
		logger.debug("---valuePart--"+valuePart);
		//���㳵��
		valuePart = calCL(valuePart);
		statePart += "," + "CL";

		String[] valueParts = valuePart.split(",");

		//����ȼ������ȡ��ȼ����������ֶ�
		result = calRllx(statePart, valuePart, insertPreFix, valueParts);
		logger.debug("---Ԥ���������ɵ�insert���---" + result);
		return result;
	}

	/**
	 * ���㳵��
	 * @param valuePart
	 * @return
	 */
	private String calCL(String valuePart) {
		// ���㳵��
		String fzDateVal = valuePart.split(",")[6].replace("'", "");// ��÷�֤���ڵ�ֵ
		try {
			Date fzDate = DateUtils.parseDate(fzDateVal, "yyyy-MM-dd");
			Date nowFullYearDate = DateUtils.ceiling(new Date(), Calendar.YEAR);
			long delta = nowFullYearDate.getTime() - fzDate.getTime();
			BigDecimal cl = new BigDecimal(delta).divide(new BigDecimal(1000))
					.divide(new BigDecimal(60), 5, BigDecimal.ROUND_HALF_UP)
					.divide(new BigDecimal(60), 5, BigDecimal.ROUND_HALF_UP)
					.divide(new BigDecimal(24), 5, BigDecimal.ROUND_HALF_UP)
					.divide(new BigDecimal(365), 1, BigDecimal.ROUND_HALF_UP);
			valuePart += "," + "'" + cl.doubleValue() + "'";
			logger.debug("---���������---"+cl);
		} catch (ParseException e) {
			logger.error(e.getMessage(), e);
		}
		return valuePart;
	}

	/**
	 * ȼ�����ͼ���
	 * @param statePart
	 * @param valuePart
	 * @param insertPreFix
	 * @param valueParts
	 * @return
	 */
	private String calRllx(String statePart, String valuePart, String insertPreFix, String[] valueParts) {
		String result;
		// ����Ȼ������ ��� rllx1����rllx2����rllx3�ֶ�
		String rllx = "";
		if (StringUtils.containsIgnoreCase(insertPreFix, "CSGJJCB")) {
			rllx = valueParts[23];
		}else if (StringUtils.containsIgnoreCase(insertPreFix, "NCKYJCB")) {
			rllx = valueParts[23];
		}else if (StringUtils.containsIgnoreCase(insertPreFix, "CZQCJCB")) {
			rllx = valueParts[23];
		}

		if ("'����'".equals(rllx)) {
			statePart = StringUtils.replace(statePart, "RLLX*1,", "");
			valuePart = StringUtils.replace(valuePart, ",'��ȼ��-����'", "");
		} else if ("'����'".equals(rllx)) {
			statePart = StringUtils.replace(statePart, "RLLX*1,", "");
			valuePart = StringUtils.replace(valuePart, ",'��ȼ��-����'", "");
		} else if ("'LPG'".equals(rllx)) {
			statePart = StringUtils.replace(statePart, "RLLX*1,", "");
			valuePart = StringUtils.replace(valuePart, ",'��ȼ��-LPG'", "");
		} else if ("'CNG'".equals(rllx)) {
			statePart = StringUtils.replace(statePart, "RLLX*1,", "");
			valuePart = StringUtils.replace(valuePart, ",'��ȼ��-CNG'", "");
		} else if ("'LNG'".equals(rllx)) {
			statePart = StringUtils.replace(statePart, "RLLX*1,", "");
			valuePart = StringUtils.replace(valuePart, ",'��ȼ��-LNG'", "");
		} else if ("'˫ȼ��'".equals(rllx)) {
			statePart = StringUtils.replace(statePart, "RLLX*1", "RLLX1" + "," + "RLLX2");
			if (StringUtils.contains(valuePart, "'����+LPG'")) {
				valuePart = StringUtils.replace(valuePart, "'����+LPG'", "'����','LPG'");
			} else if (StringUtils.contains(valuePart, "'����+CNG'")) {
				valuePart = StringUtils.replace(valuePart, "'����+CNG'", "'����','CNG'");
			} else if (StringUtils.contains(valuePart, "'����+LNG'")) {
				valuePart = StringUtils.replace(valuePart, "'����+LNG'", "'����','LNG'");
			} else if (StringUtils.contains(valuePart, "'����+LPG'")) {
				valuePart = StringUtils.replace(valuePart, "'����+LPG'", "'����','LPG'");
			} else if (StringUtils.contains(valuePart, "'����+CNG'")) {
				valuePart = StringUtils.replace(valuePart, "'����+CNG'", "'����','CNG'");
			} else if (StringUtils.contains(valuePart, "'����+LNG'")) {
				valuePart = StringUtils.replace(valuePart, "'����+LNG'", "'����','LNG'");
			}
		} else if ("'��϶���'".equals(rllx)) {
			statePart = StringUtils.replace(statePart, "RLLX*1", "RLLX3");
		}

		// �������
		result = insertPreFix + statePart + ExcelConstants.SQL_INSERT_VALUE_FLAG + valuePart + ")";
		return result;
	}

}
