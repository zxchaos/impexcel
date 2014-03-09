package com.taiji.excelimp.core;

import java.io.File;
import java.math.BigDecimal;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Properties;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.dom4j.Document;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.taiji.excelimp.db.DBAccess;
import com.taiji.excelimp.util.ExcelConstants;
import com.taiji.excelimp.util.ExcelImportUtil;

public class DataBatchImpExcel extends AbstractImpExcel{
	public static Logger logger = LoggerFactory.getLogger(DataBatchImpExcel.class);

	/**
	 * ����excelDirĿ¼�µ�excel�ļ�
	 * 
	 * @param excelDir
	 */
	public void importExcel(Properties sysConfig, DBAccess dbAccess) throws Exception {
		logger.info("+++�����������ܵ��빦�ܿ�ʼ+++");
		String impDir = sysConfig.getProperty("impDir")+File.separator+sysConfig.getProperty("dataBatchImpDirName")+File.separator;
		String configFilePath = sysConfig.getProperty("configFilePath");
		String backupDir = sysConfig.getProperty("backupDir") + File.separator + getNowDateString() + File.separator;

		File[] excelFiles = checkImpDir(impDir);
		if (excelFiles == null) {
			return;
		}

		for (File excelFile : excelFiles) {
			logger.debug("---��ʼ�����ļ�---" + excelFile.getAbsolutePath());
			Map<String, String> resultMap = new HashMap<String, String>();// ��Ž��������Map
			// �ϴ��� ftp
			// ��excel���ļ�����������Ϊ��uuid_�������_ʡ����_�д���_�ش���_��λid_���_�·�_ʡ����_������_������_��λ����.xlsx�����в�������Ӧ
			// templateId ����
			// ���в��������������й�����csgj��,����������czqc��,ũ����ˣ�ncky��
			String fileName = excelFile.getName().substring(0, excelFile.getName().lastIndexOf("."));
			logger.debug("ȥ����չ������ļ���---" + fileName);
			String[] fileNameParts = fileName.split("_");
			// ��ȡ�������
			String templateId = fileNameParts[1];
			Workbook workbook = null;
			Document document = null;
			long dwid = Long.valueOf(fileNameParts[5]);
			String hylb = fileNameParts[1];
			Map<String, String>infoFieldMap = new HashMap<String, String>();
			infoFieldMap.put("hylb", hylb);
			infoFieldMap.put("dwid", String.valueOf(dwid));
			boolean isSuccess = true;
			try {
				// ��������ļ��ĵ�������Ҫ�������ļ��Ĺ�����������ͬʱ��֤�ļ��Ƿ�Ϊϵͳ�ṩģ��
				logger.debug("---��ʼ���ģ��У���ַ�������������ļ��ĵ������뵼��excel�ļ��Ĺ���������---");
				document = ExcelImportUtil.getConfigFileDoc(configFilePath);
				workbook = ExcelImportUtil.genWorkbook(excelFile, document, templateId, resultMap);
				logger.debug("---�������ģ��У���ַ�������������ļ��ĵ������뵼��excel�ļ��Ĺ���������---");

				if (ExcelConstants.SUCCESS.equals(resultMap.get(ExcelConstants.RESULT_KEY))) {
					// Ԥ�ȹ�����
					// excel �ļ���Ԥ����֤�ò�����ҵ����أ� ���������������ж� ���� ���� cng lng lpg ����д���
					preCheck(sysConfig, templateId, resultMap, workbook);// Ԥ�ȼ���
					if (ExcelConstants.FAIL.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
						ExcelImportUtil.setFailMsg(resultMap, "ȼ������У��", false);
						ExcelImportUtil.setFailMsg(resultMap, "��ֵ����У��");
					}
					ExcelImportUtil.importExcel(workbook, document, templateId, resultMap);// ����updatesql
				}
				
				String insertSqls = resultMap.get(ExcelConstants.SQLS_KEY);
				if (!ExcelConstants.FAIL.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY)) && StringUtils.isBlank(insertSqls)) {
					logger.info("+++���ɵ�sqlΪ��+++������ģ����û������");
					ExcelImportUtil.setFailMsg(resultMap, "�����ģ���в���������");
					isSuccess = false;
				}
				
				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					String[] updateSqls = remakeUpdateSql(resultMap, excelFile, fileNameParts);
					logger.info("---�ļ�---" + excelFile.getName() + "---�����ɹ�---");
					
					//�������º͵��ô洢����
					dbAccess.updateAndCallprocedure(updateSqls, fileNameParts);
					super.insertImpInfo(dbAccess, resultMap, infoFieldMap, true, super.getType());
					logger.debug("---�ɹ���Ϣ�������---");
				}

				if (ExcelConstants.FAIL.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					// ��������Ϣд����־
					logger.debug("---������Ϣд�����ݿ�---");
					super.insertImpInfo(dbAccess, resultMap, infoFieldMap, false, super.getType());
					logger.debug("---������Ϣд�����---");
					isSuccess = false;
				}
			} catch (Exception e) {
				isSuccess = false;
				logger.info("+++��������쳣+++");
				logger.error(e.getMessage(), e);
				resultMap.remove(ExcelConstants.MSG_KEY);
				ExcelImportUtil.setFailMsg(resultMap, "�����쳣,����ϵ������Ա");
				super.insertImpInfo(dbAccess, resultMap, infoFieldMap, false, super.getType());
				logger.info("+++����������һ���ļ�+++");
				continue;
			}finally{
				// �����ļ�
				backupFile(excelFile, backupDir, isSuccess);
			}
		}
	}

	/**
	 * �������update���
	 * 
	 * @param resultMap
	 * @param excelFile
	 * @param fileNameParts
	 */
	private  String[] remakeUpdateSql(Map<String, String> resultMap, File excelFile, String[] fileNameParts) {
		// ÿ���ļ�����ɹ��� ƴװ�������� update ���
		String[] updateSqls = resultMap.get(ExcelConstants.SQLS_KEY).split(ExcelConstants.SQL_TAIL);

		logger.debug("---�ļ�---" + excelFile.getName() + "�����ɹ�---���ɵ�update��䣺");
		if (logger.isDebugEnabled()) {
			for (int i = 0; i < updateSqls.length; i++) {
				logger.debug(updateSqls[i]);
			}
		}
		for (int i = 0; i < updateSqls.length; i++) {
			Map<String, String> updateMap = updateToMap(updateSqls[i]);
			// ���㵱����ʻ���
			// Double dyxslc = calDYXSLC(updateMap);
			// �����վ���ʻ���
			calRJXSLC(updateMap);
			// ����ٹ���ƽ������
			calBGLPJDH(updateMap);
			// �������update���
			updateSqls[i] = makeUpdateSql(updateSqls[i], updateMap, fileNameParts[6], fileNameParts[7],
					fileNameParts[5]);
		}
		if (logger.isDebugEnabled()) {
			logger.debug("---�������ɵ�update���---");
			for (int i = 0; i < updateSqls.length; i++) {
				logger.debug(updateSqls[i]);
			}
		}
		return updateSqls;
	}

	/**
	 * ��������Ϣд����־
	 * 
	 * @param dbAccess
	 * @param resultMap
	 * @param fileNameParts
	 */
	public  void insertImpInfo(DBAccess dbAccess, Map<String, String> resultMap, String[] fileNameParts,
			boolean isSuccess) throws Exception{
		logger.debug("---����������Ϣд�����ݿ�---");
		long dwid = Long.valueOf(fileNameParts[5]);
		String hylb = fileNameParts[1];
		String sj = getNowDateString("yyyy��MM��dd��HHʱmm��");
		String info = "";
		if (isSuccess) {
			info = "���������Ϣ�ɹ�";
		} else {
			info = resultMap.get(ExcelConstants.MSG_KEY);
		}
		logger.info("+++��������Ϣ+++\n" + info);
		dbAccess.insertErrorInfo(dwid, hylb, sj, info, super.getType());
		logger.debug("---��������Ϣд�����ݿ����---");
	}


	/**
	 * ���ݽ������updateMap�еĵ�����ʻ���ֵ�����վ���ʻ���
	 * 
	 * @param updateMap
	 */
	private static void calRJXSLC(Map<String, String> updateMap) {
		Double dyxslc = Double.valueOf(updateMap.get("DYXSLC"));
		calRJXSLC(dyxslc, updateMap);
	}

	/**
	 * ���ݽ������update����е�ֵ�͵�����ʻ��̼����վ���ʻ���
	 * 
	 * @param dyxslc
	 * @param updateMap
	 */
	private static void calRJXSLC(Double dyxslc, Map<String, String> updateMap) {
		BigDecimal sjyytsBD = new BigDecimal(updateMap.get("SJYYTS"));
		BigDecimal dyxslcBD = BigDecimal.valueOf(dyxslc);
		if (sjyytsBD.doubleValue() != 0) {
			BigDecimal rjxslcBD = dyxslcBD.divide(sjyytsBD, 2, BigDecimal.ROUND_HALF_UP);
			updateMap.put("RJXSLC", String.valueOf(rjxslcBD.doubleValue()));
		}else {
			updateMap.put("RJXSLC", "0");
		}
	}

	/**
	 * �������update���
	 * 
	 * @param origUpdateSql
	 * @param updateMap
	 */
	private static String makeUpdateSql(String origUpdateSql, Map<String, String> updateMap, String year, String month,
			String dwid) {
		String newUpdate = "";
		// ��ȡupdateͷ��: UPDATE TABLENAME SET
		String updateHead = origUpdateSql.substring(0, origUpdateSql.lastIndexOf("SET") + 3);
		// ��ȡwhere�Ӿ�
		StringBuffer updateWhere = new StringBuffer(origUpdateSql.substring(origUpdateSql.lastIndexOf("WHERE")));
		// update������������� �·� ��λid��������
		updateWhere.append(" AND NF=");
		updateWhere.append(year);
		updateWhere.append(" AND YF=");
		updateWhere.append(month);
		updateWhere.append(" AND DWID=");
		updateWhere.append(dwid);
		// ���ӹ̶�д�� ��������Ϊ����
		updateWhere.append(" AND BGQK<>'����' ");

		// ȥ��ȼ�������ֶν�����map�е�key-value ����Ϊupdate�е�set����
		StringBuffer setPart = new StringBuffer("");
		for (Entry<String, String> entry : updateMap.entrySet()) {
			if ("RLLX".equals(entry.getKey())) {// ���������ɵ�update�г�ȥȼ�������ֶ�
				continue;
			}
			setPart.append(entry.getKey());
			setPart.append("=");
			setPart.append(entry.getValue());
			setPart.append(",");
		}
		String sPart = setPart.substring(0, setPart.lastIndexOf(","));
		// ƴ��update
		newUpdate = updateHead + " " + sPart + " " + updateWhere;
		return newUpdate;
	}

	/**
	 * ���㵱����ʻ���
	 * 
	 * @param updateMap
	 * @return
	 */
	private static Double calDYXSLC(Map<String, String> updateMap) {
		// �����³���ʻ��̺���ĩ��ʻ��̼��� DYXSLC ������ʻ����
		Double ycgl = new Double(updateMap.get("YCGL"));// �³���ʻ���
		Double ymgl = new Double(updateMap.get("YMGL"));// ��ĩ��ʻ���
		Double dyxslc = ymgl - ycgl;// ������ʻ���=��ĩ��ʻ���-�³���ʻ���
		updateMap.put("DYXSLC", dyxslc.toString());
		return dyxslc;
	}

	/**
	 * ��update sql�е�set֮���Ҫ���µ��ֶκ�ֵ�ŵ�map�� update �����ʽΪ UPDATE TABLENAME SET
	 * FIELD1=VALUE1,FIELD2=VALUE2...
	 * 
	 * @param updateSql
	 * @return
	 */
	 @Override
	public  Map<String, String> updateToMap(String updateSql) {
		// logger.debug("---��update����е�set���ַ��õ�map��----");
		Map<String, String> resultMap = new HashMap<String, String>();
		String convertUpdate = updateSql.substring(updateSql.lastIndexOf("SET") + 3).replace(ExcelConstants.SQL_TAIL, "").trim();
		// logger.debug("---ȥ��update���ͷ��---" + convertUpdate);
		// ȥ��WHERE�Ӿ�
		convertUpdate = convertUpdate.substring(0, convertUpdate.lastIndexOf("WHERE")).trim();
		// logger.debug("---ȥ��where�Ӿ��---" + convertUpdate);
		String[] fieldsValues = convertUpdate.split(",");
		for (String fv : fieldsValues) {
			String[] updatedFieldValue = fv.split("=");
			resultMap.put(updatedFieldValue[0], updatedFieldValue[1]);
		}
		// ���ӹ̶�д�� wczt = 1
		resultMap.put("WCZT", "1");
		// logger.debug("---���ɵ�map---" + resultMap);
		return resultMap;
	}

	/**
	 * ����ҵ��Ԥ�ȼ���
	 * 
	 * @param sysConfig
	 *            ϵͳ����
	 * @param type
	 * @param resultMap
	 *            ��ż�����
	 * @return
	 */
	public void preCheck(Properties sysConfig, String type, Map<String, String> resultMap, Workbook workbook)
			throws Exception {
		logger.debug("---����---Ԥ�ȹ������---");
		Integer dataRowStartNum = Integer.valueOf(sysConfig.getProperty(type + "DataRowStartNum"));
		String ruleString = "";// �����ַ���
		logger.debug("---�����������---" + type);
		String rulesProp = type + "CheckRules";
		ruleString = sysConfig.getProperty(rulesProp);
		checkRule(ruleString, resultMap, type, dataRowStartNum, workbook);
		logger.debug("---����----Ԥ�ȹ������---");
	}

	/**
	 * ������
	 * 
	 * @param ruleString
	 *            �����ڣ�11-���ȼ��:(22|23)&(24|25|26);11-˫ȼ��:(22|23) ��ʾ���������;����
	 * @param excelFile
	 * @param resultMap
	 * @param type
	 */
	private static void checkRule(String ruleString, Map<String, String> resultMap, String type,
			Integer dataStartRowNum, Workbook workbook) throws Exception {
		String[] checkRules = ruleString.split(";");
		Sheet sheet = workbook.getSheet(type);
		for (int i = dataStartRowNum - 1; i <= sheet.getLastRowNum(); i++) {
			checkRowRules(resultMap, checkRules, sheet, i);
			if (ExcelConstants.FAIL.equals(resultMap.get(ExcelConstants.RESULT_KEY))) {
				String failMsg = resultMap.get(ExcelConstants.MSG_KEY);
				logger.warn("������ʧ�ܣ�" + failMsg);
			}
		}
	}

	/**
	 * ��ÿһ��Ӧ�ü�����
	 * 
	 * @param resultMap
	 * @param checkRules
	 * @param sheet
	 * @param rowIndex
	 */
	private static void checkRowRules(Map<String, String> resultMap, String[] checkRules, Sheet sheet, int rowIndex) {
		logger.debug("---Ԥ�ȼ�����к�---"+rowIndex);
		Row row = sheet.getRow(rowIndex);
		if (row == null) {
			logger.warn("+++�У�"+(rowIndex+1)+"---Ϊ��+++");
			return;
		}else {
			Cell cphmCell = row.getCell(0);//���ÿһ�еĵ�һ�м����ƺ���
			if (StringUtils.isBlank(ExcelImportUtil.getCellValue(cphmCell))) {
				logger.debug("+++�У�"+(rowIndex+1)+"---�ĳ��ƺ�����Ϊ�գ���������������");
				return;
			}
		}
		
		for (int i = 0; i < checkRules.length; i++) {// ��֤����
			String[] ruleCell = checkRules[i].split(":");
			String[] typeCell = ruleCell[0].split("-");// ruleCell[0]Ϊ���Ͳ��֣����磺11-���ȼ��
			Integer colIndex = Integer.valueOf(typeCell[0]) - 1;// typeCell[0]Ϊ���Ͳ��ֵ��к�
			Cell tCell = row.getCell(colIndex);// ���Ͳ��ֵ�Ԫ��
			if (tCell != null && StringUtils.isNotBlank(ExcelImportUtil.getCellValue(tCell))) {
				String cellValue = ExcelImportUtil.getCellValue(tCell);

				if (cellValue.equals(typeCell[1])) {// ����Ԫ���ֵ�����õ�����ֵ��typeCell[1]��������������У��
					if (ruleCell[1].contains("@")) {// @������кŴ�����в�����ֵ
						String nonIncludeCols = ruleCell[1].split("@")[1];

						if (nonIncludeCols.contains("(")) {// ȥ������
							nonIncludeCols = nonIncludeCols.replace("(", "").replace(")", "");
						}
						checkNonIncludeCols(resultMap, rowIndex, row, nonIncludeCols);

						String ruleCheckCols = ruleCell[1].split("@")[0];
						ruleCell[1] = ruleCheckCols;
					}

					if (ExcelConstants.FAIL.equals(resultMap.get(ExcelConstants.RESULT_KEY))) {
						logger.warn("������ʧ��");
						break;
					}

					if (ruleCell[1].contains("&")) {
						// ruleCell[1]Ϊ�������֣����а��������ҡ��ṹ
						String[] conditions = ruleCell[1].split("&");

						for (int j = 0; j < conditions.length; j++) {
							if (conditions[j].contains("(")) {// ������()
																// ˵��()�а�������򡱽ṹ
								String condition = conditions[j].replace("(", "").replace(")", "");// ȥ������
								checkOr(resultMap, row, condition);
								if (ExcelConstants.FAIL.equals(resultMap.get(ExcelConstants.RESULT_KEY))) {
									logger.warn(resultMap.get(ExcelConstants.MSG_KEY));
									break;
								}
							} else {// ��������������Ϊ��ֵ
								String nullCellNum = "";// ��Ԫ��Ϊ�յ��к��ö��ŷָ�
								String colNum = conditions[j];
								Cell cCell = row.getCell(Integer.valueOf(colNum) - 1);
								if (cCell == null || StringUtils.isBlank(ExcelImportUtil.getCellValue(cCell))) {
									nullCellNum += ExcelImportUtil.colNumToColName(Integer.valueOf(colNum)) + ",";
									ExcelImportUtil.setFailMsg(resultMap, "����ȼ�������жϣ��У�" + (rowIndex + 1) + "���У�"
											+ nullCellNum + "�ж�û��ֵ");
									logger.warn(resultMap.get(ExcelConstants.MSG_KEY));
									break;
								}
							}
						}
					} else {// �����������ҡ��ṹ ��ôֻ�� ���򡱽ṹ���ߵ���
						if (ruleCell[1].contains("|")) {
							checkOr(resultMap, row, ruleCell[1]);
						} else {// ���������򡱽ṹ ֻ�е���
							Integer cellIndex = Integer.valueOf(ruleCell[1]) - 1;
							Cell cCell = row.getCell(Integer.valueOf(ruleCell[1]) - 1);
							if (cCell == null || StringUtils.isBlank(ExcelImportUtil.getCellValue(cCell))) {
								String failMsg = ExcelImportUtil.makeFailMsg(rowIndex, cellIndex, "������ֵ");
								ExcelImportUtil.setFailMsg(resultMap, "����ȼ�������жϣ�" + failMsg);
								logger.warn(resultMap.get(ExcelConstants.MSG_KEY));
							}
						}
					}
					break;
				}
			} else {
				String failMsg = ExcelImportUtil.makeFailMsg(rowIndex, colIndex, "������ֵ");
				ExcelImportUtil.setFailMsg(resultMap, failMsg);
				logger.warn(resultMap.get(ExcelConstants.MSG_KEY));
				break;
			}
		}
	}

	/**
	 * @param resultMap
	 * @param rowIndex
	 * @param row
	 * @param ruleCell
	 */
	private static void checkNonIncludeCols(Map<String, String> resultMap, int rowIndex, Row row, String nonIncludeCols) {
		if (nonIncludeCols.contains(",")) {
			String nonCellNum = "";
			String[] nonInc = nonIncludeCols.split(",");
			StringBuffer failMsg = new StringBuffer("");
			boolean isFail = false;
			for (int j = 0; j < nonInc.length; j++) {
				// logger.debug("---����---checkNonIncludeCols---���������к�---" +
				// nonInc[j]);

				Cell cCell = row.getCell(Integer.valueOf(nonInc[j]) - 1);

				if (cCell != null) {
					boolean isNull = true;
					switch (cCell.getCellType()) {
					case Cell.CELL_TYPE_BLANK:
						isNull = true;
						break;
					case Cell.CELL_TYPE_STRING:
						isNull = StringUtils.isBlank(cCell.getRichStringCellValue().getString());
						break;
					case Cell.CELL_TYPE_NUMERIC:
						isNull = StringUtils.isBlank(String.valueOf(cCell.getNumericCellValue()));
						break;
					default:
						break;
					}
					if (!isNull) {
						nonCellNum += ExcelImportUtil.colNumToColName(Integer.valueOf(nonInc[j])) + ",";
						isFail = true;
					}

				}
			}
			if (isFail) {
				failMsg.append("�������������жϣ��У�" + (rowIndex + 1) + "���У�" + nonCellNum + "�в�����ֵ");
				ExcelImportUtil.setFailMsg(resultMap, failMsg.toString());
				logger.warn(resultMap.get(ExcelConstants.MSG_KEY));
			}

		} else {
			Cell cCell = row.getCell(Integer.valueOf(nonIncludeCols) - 1);
			if (cCell != null) {
				ExcelImportUtil
						.setFailMsg(resultMap, "�������������жϣ��У�" + (rowIndex + 1) + "���У�" + nonIncludeCols + "�в�����ֵ");
				logger.warn(resultMap.get(ExcelConstants.MSG_KEY));
			}
		}
	}

	/**
	 * ��֤�����еġ��򡱽ṹ����֤
	 * 
	 * @param resultMap
	 * @param rowIndex
	 * @param row
	 * @param nullCellNum
	 * @param condition
	 */
	private static void checkOr(Map<String, String> resultMap, Row row, String condition) {
		String nullCellNum = "";// ��Ԫ��Ϊ�յ��к��ö��ŷָ�
		String[] colNum = condition.split("\\|");// ������򡱽ṹ�ָ�Ϊ�к�
		String notNullCellNum = "";// ��Ԫ��Ϊ�յ��к��ö��ŷָ�
		int orCount = 0;// ���򡱽ṹ�в�Ϊ�յĵ�Ԫ��ĸ���
		for (int k = 0; k < colNum.length; k++) {
			Cell cCell = row.getCell(Integer.valueOf(colNum[k]) - 1);
			if (cCell != null && StringUtils.isNotBlank(ExcelImportUtil.getCellValue(cCell))) {
				orCount++;
				notNullCellNum += ExcelImportUtil.colNumToColName(Integer.valueOf(colNum[k])) + "��";
			} else {
				nullCellNum += ExcelImportUtil.colNumToColName(Integer.valueOf(colNum[k])) + "��";
			}
		}
		if (orCount > 1) {// ��ṹ��ֻ����1�е���ֵ
			notNullCellNum = StringUtils.substring(notNullCellNum, 0, StringUtils.lastIndexOf(notNullCellNum, "��"));
			ExcelImportUtil.setFailMsg(resultMap, "����ȼ�������жϣ��У�" + (row.getRowNum() + 1) + "���У�" + notNullCellNum
					+ "��ֻ����һ����ֵ");
			logger.warn(resultMap.get(ExcelConstants.MSG_KEY));
		} else if (orCount == 0) {
			nullCellNum = StringUtils.substring(nullCellNum, 0, StringUtils.lastIndexOf(nullCellNum, "��"));
			ExcelImportUtil.setFailMsg(resultMap, "����ȼ�������жϣ��У�" + (row.getRowNum() + 1) + "���У�" + nullCellNum
					+ "�б�����һ����ֵ");
			logger.warn(resultMap.get(ExcelConstants.MSG_KEY));
		}
	}

	


	/**
	 * �����ȼ�ϵİٹ���ƽ������
	 * 
	 * @param updateMap
	 *            ����ʱȡ�õĸ��ֶε�ֵ��Դ���Լ�������Ϻ��ȼ�ϵİٹ���ƽ�����ı����map
	 */
	private static void calBGLPJDH(Map<String, String> updateMap) {
		Double dyxslc = Double.valueOf(updateMap.get("DYXSLC"));
		calBGLPJDH(dyxslc, updateMap);
	}

	/**
	 * ����ȼ�����ͼ����ȼ�ϵİٹ���ƽ������
	 * 
	 * @param dyxslc
	 *            ������ʻ���
	 * @param updateMap
	 *            ����ʱȡ�õĸ��ֶε�ֵ��Դ���Լ�������Ϻ��ȼ�ϵİٹ���ƽ�����ı����map
	 */
	private static void calBGLPJDH(Double dyxslc, Map<String, String> updateMap) {
		// logger.debug("---��ʼ����ٹ���ƽ������---");
		String rllx = updateMap.get("RLLX");// ȡ��ȼ������
		// logger.debug("---ȼ������---" + rllx);
		if ("'����'".equals(rllx)) {
			calBGLPJDHNormal("DYQYXHL", "QYBGLPJDH", dyxslc, updateMap);
		} else if ("'����'".equals(rllx)) {
			calBGLPJDHNormal("DYCYXHL", "CYBGLPJDH", dyxslc, updateMap);
		} else if ("'LNG'".equals(rllx)) {
			calBGLPJDHNormal("DYLNGXHL", "LNGBGLPJDH", dyxslc, updateMap);
		} else if ("'LPG'".equals(rllx)) {
			calBGLPJDHNormal("DYLPGXHL", "LPGBGLPJDH", dyxslc, updateMap);
		} else if ("'CNG'".equals(rllx)) {
			calBGLPJDHNormal("DYCNGXHL", "CNGBGLPJDH", dyxslc, updateMap);
		} else if ("'˫ȼ��'".equals(rllx)) {
			// �ӵ���������������DYQYXHL�����²�����������DYCYXHL����ȡһ���ǿ�ֵ����updateMap�� �ٴ�
			// LNG��LPG��CNG�ĵ��²�����������ȡһ���ǿ�ֵ����updateMap��
			calBGLPJDHNormal("DYQYXHL", "QYBGLPJDH", dyxslc, updateMap);
			calBGLPJDHNormal("DYCYXHL", "CYBGLPJDH", dyxslc, updateMap);
			calBGLPJDHNormal("DYLNGXHL", "LNGBGLPJDH", dyxslc, updateMap);
			calBGLPJDHNormal("DYLPGXHL", "LPGBGLPJDH", dyxslc, updateMap);
			calBGLPJDHNormal("DYCNGXHL", "CNGBGLPJDH", dyxslc, updateMap);
		} else if ("���ȼ��".equals(rllx)) {
			// �ӵ���������������DYQYXHL�����²�����������DYCYXHL����ȡһ���ǿ�ֵ����updateMap��
			calBGLPJDHNormal("DYQYXHL", "QYBGLPJDH", dyxslc, updateMap);
			calBGLPJDHNormal("DYCYXHL", "CYBGLPJDH", dyxslc, updateMap);
		}
		// logger.debug("---��������ٹ���ƽ������---");
	}

	/**
	 * ����ȼ�����ͣ����ͣ����ͣ�LNG��LPG,CNG������ٹ���ƽ������
	 * 
	 * @param dyxhl
	 *            ��ȡ�����������ֶ�����
	 * @param bglpjdh
	 *            �洢�ٹ���ƽ�������ֶ�����
	 * @param updateMap
	 *            ������ʱ��������Դ�������Ϻ�����ݴ洢
	 */
	private static void calBGLPJDHNormal(String dyxhl, String bglpjdh, Double dyxslc, Map<String, String> updateMap) {
		if (updateMap.get(dyxhl) != null && StringUtils.isNotBlank(updateMap.get(dyxhl))) {
			BigDecimal dyxhlBD = new BigDecimal(updateMap.get(dyxhl));// ����������
			BigDecimal dyxslcBD = BigDecimal.valueOf(dyxslc);
			if (dyxslcBD.doubleValue() != 0) {
				// �ٹ���ƽ������ = ����������/������ʻ���� * 100
				BigDecimal bglpjdhBD = dyxhlBD.divide(dyxslcBD, 4, BigDecimal.ROUND_HALF_UP).multiply(
						BigDecimal.valueOf(100));
				updateMap.put(bglpjdh, String.valueOf(bglpjdhBD.doubleValue()));// �ٹ���ƽ������
			}else {
				updateMap.put(bglpjdh, "0");// �ٹ���ƽ������
			}
			
		}
	}

	
}
