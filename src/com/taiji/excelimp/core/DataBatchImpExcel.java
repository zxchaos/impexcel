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
	 * 导入excelDir目录下的excel文件
	 *
	 * @param excelDir
	 */
	public void importExcel(Properties sysConfig, DBAccess dbAccess) throws Exception {
		logger.info("+++数据批量功能导入功能开始+++");
		String impDir = sysConfig.getProperty("impDir")+File.separator+sysConfig.getProperty("dataBatchImpDirName")+File.separator;
		String configFilePath = sysConfig.getProperty("configFilePath");
		String backupDir = sysConfig.getProperty("backupDir") + File.separator + getNowDateString() + File.separator;

		File[] excelFiles = checkImpDir(impDir);
		if (excelFiles == null) {
			return;
		}

		for (File excelFile : excelFiles) {
			logger.debug("---开始解析文件---" + excelFile.getAbsolutePath());
			Map<String, String> resultMap = new HashMap<String, String>();// 存放解析结果的Map
			// 上传到 ftp
			// 的excel的文件的命名规则为：uuid_操作类别_省代码_市代码_县代码_单位id_年份_月份_省名称_市名称_县名称_单位名称.xlsx；其中操作类别对应
			// templateId 参数
			// 其中操作类别包括：城市公交（csgj）,出租汽车（czqc）,农村客运（ncky）
			String fileName = excelFile.getName().substring(0, excelFile.getName().lastIndexOf("."));
			logger.debug("去掉扩展名后的文件名---" + fileName);
			String[] fileNameParts = fileName.split("_");
			// 获取操作类别
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
				// 获得配置文件文档对象与要解析的文件的工作簿对象获得同时验证文件是否为系统提供模板
				logger.debug("---开始检查模板校验字符串并获得配置文件文档对象与导入excel文件的工作簿对象---");
				document = ExcelImportUtil.getConfigFileDoc(configFilePath);
				workbook = ExcelImportUtil.genWorkbook(excelFile, document, templateId, resultMap);
				logger.debug("---结束检查模板校验字符串并获得配置文件文档对象与导入excel文件的工作簿对象---");

				if (ExcelConstants.SUCCESS.equals(resultMap.get(ExcelConstants.RESULT_KEY))) {
					// 预先规则检查
					// excel 文件的预先验证该步骤与业务相关： 根据油料类型来判断 汽油 柴油 cng lng lpg 的填写情况
					preCheck(sysConfig, templateId, resultMap, workbook);// 预先检验
					if (ExcelConstants.FAIL.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
						ExcelImportUtil.setFailMsg(resultMap, "燃料类型校验", false);
						ExcelImportUtil.setFailMsg(resultMap, "数值类型校验");
					}
					ExcelImportUtil.importExcel(workbook, document, templateId, resultMap);// 生成updatesql
				}

				String insertSqls = resultMap.get(ExcelConstants.SQLS_KEY);
				if (!ExcelConstants.FAIL.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY)) && StringUtils.isBlank(insertSqls)) {
					logger.info("+++生成的sql为空+++可能是模板中没有数据");
					ExcelImportUtil.setFailMsg(resultMap, "导入的模板中不包含数据");
					isSuccess = false;
				}

				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					String[] updateSqls = remakeUpdateSql(resultMap, excelFile, fileNameParts);
					logger.info("---文件---" + excelFile.getName() + "---解析成功---");

					//批量更新和调用存储过程
					dbAccess.updateAndCallprocedure(updateSqls, fileNameParts,hylb);
					super.insertImpInfo(dbAccess, resultMap, infoFieldMap, true, super.getType());
					logger.debug("---成功信息插入完毕---");
				}

				if (ExcelConstants.FAIL.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					// 将出错信息写入日志
					logger.debug("---出错信息写入数据库---");
					super.insertImpInfo(dbAccess, resultMap, infoFieldMap, false, super.getType());
					logger.debug("---出错信息写入完毕---");
					isSuccess = false;
				}
			} catch (Exception e) {
				isSuccess = false;
				logger.info("+++导入出现异常+++");
				logger.error(e.getMessage(), e);
				resultMap.remove(ExcelConstants.MSG_KEY);
				ExcelImportUtil.setFailMsg(resultMap, "导入异常,请联系管理人员");
				super.insertImpInfo(dbAccess, resultMap, infoFieldMap, false, super.getType());
				logger.info("+++继续导入下一个文件+++");
				continue;
			}finally{
				// 备份文件
				backupFile(excelFile, backupDir, isSuccess);
			}
		}
	}

	/**
	 * 重新组成update语句
	 *
	 * @param resultMap
	 * @param excelFile
	 * @param fileNameParts
	 */
	private  String[] remakeUpdateSql(Map<String, String> resultMap, File excelFile, String[] fileNameParts) {
		// 每个文件导入成功后 拼装成完整的 update 语句
		String[] updateSqls = resultMap.get(ExcelConstants.SQLS_KEY).split(ExcelConstants.SQL_TAIL);

		logger.debug("---文件---" + excelFile.getName() + "解析成功---生成的update语句：");
		if (logger.isDebugEnabled()) {
			for (int i = 0; i < updateSqls.length; i++) {
				logger.debug(updateSqls[i]);
			}
		}
		for (int i = 0; i < updateSqls.length; i++) {
			Map<String, String> updateMap = updateToMap(updateSqls[i]);
			// 计算当月行驶里程
			// Double dyxslc = calDYXSLC(updateMap);
			// 计算日均行驶里程
			calRJXSLC(updateMap);
			// 计算百公里平均单耗
			calBGLPJDH(updateMap);
			// 重新组成update语句
			updateSqls[i] = makeUpdateSql(updateSqls[i], updateMap, fileNameParts[6], fileNameParts[7],
					fileNameParts[5]);
		}
		if (logger.isDebugEnabled()) {
			logger.debug("---最终生成的update语句---");
			for (int i = 0; i < updateSqls.length; i++) {
				logger.debug(updateSqls[i]);
			}
		}
		return updateSqls;
	}

	/**
	 * 将错误信息写入日志
	 *
	 * @param dbAccess
	 * @param resultMap
	 * @param fileNameParts
	 */
	public  void insertImpInfo(DBAccess dbAccess, Map<String, String> resultMap, String[] fileNameParts,
			boolean isSuccess) throws Exception{
		logger.debug("---将导入结果信息写入数据库---");
		long dwid = Long.valueOf(fileNameParts[5]);
		String hylb = fileNameParts[1];
		String sj = getNowDateString("yyyy年MM月dd日HH时mm分");
		String info = "";
		if (isSuccess) {
			info = "您导入的信息成功";
		} else {
			info = resultMap.get(ExcelConstants.MSG_KEY);
		}
		logger.info("+++导入结果信息+++\n" + info);
		dbAccess.insertErrorInfo(dwid, hylb, sj, info, super.getType());
		logger.debug("---导入结果信息写入数据库完成---");
	}


	/**
	 * 根据解析后的updateMap中的当月行驶里程值计算日均行驶里程
	 *
	 * @param updateMap
	 */
	private static void calRJXSLC(Map<String, String> updateMap) {
		Double dyxslc = Double.valueOf(updateMap.get("DYXSLC"));
		calRJXSLC(dyxslc, updateMap);
	}

	/**
	 * 根据解析后的update语句中的值和当月行驶里程计算日均行驶里程
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
	 * 重新组合update语句
	 *
	 * @param origUpdateSql
	 * @param updateMap
	 */
	private static String makeUpdateSql(String origUpdateSql, Map<String, String> updateMap, String year, String month,
			String dwid) {
		String newUpdate = "";
		// 获取update头部: UPDATE TABLENAME SET
		String updateHead = origUpdateSql.substring(0, origUpdateSql.lastIndexOf("SET") + 3);
		// 获取where子句
		StringBuffer updateWhere = new StringBuffer(origUpdateSql.substring(origUpdateSql.lastIndexOf("WHERE")));
		// update条件中增加年份 月份 单位id更新条件
		updateWhere.append(" AND NF=");
		updateWhere.append(year);
		updateWhere.append(" AND YF=");
		updateWhere.append(month);
		updateWhere.append(" AND DWID=");
		updateWhere.append(dwid);
		// 增加固定写法 变更情况不为报废
		updateWhere.append(" AND BGQK<>'报废' ");

		// 去掉燃料类型字段将其余map中的key-value 整合为update中的set部分
		StringBuffer setPart = new StringBuffer("");
		for (Entry<String, String> entry : updateMap.entrySet()) {
			if ("RLLX".equals(entry.getKey())) {// 在重新生成的update中出去燃料类型字段
				continue;
			}
			setPart.append(entry.getKey());
			setPart.append("=");
			setPart.append(entry.getValue());
			setPart.append(",");
		}
		String sPart = setPart.substring(0, setPart.lastIndexOf(","));
		// 拼接update
		newUpdate = updateHead + " " + sPart + " " + updateWhere;
		return newUpdate;
	}

	/**
	 * 计算当月行驶里程
	 *
	 * @param updateMap
	 * @return
	 */
	private static Double calDYXSLC(Map<String, String> updateMap) {
		// 根据月初行驶里程和月末行驶里程计算 DYXSLC 当月行驶历程
		Double ycgl = new Double(updateMap.get("YCGL"));// 月初行驶里程
		Double ymgl = new Double(updateMap.get("YMGL"));// 月末行驶里程
		Double dyxslc = ymgl - ycgl;// 当月行驶里程=月末行驶里程-月初行驶里程
		updateMap.put("DYXSLC", dyxslc.toString());
		return dyxslc;
	}

	/**
	 * 将update sql中的set之后的要更新的字段和值放到map中 update 语句形式为 UPDATE TABLENAME SET
	 * FIELD1=VALUE1,FIELD2=VALUE2...
	 *
	 * @param updateSql
	 * @return
	 */
	 @Override
	public  Map<String, String> updateToMap(String updateSql) {
		// logger.debug("---将update语句中的set部分放置到map中----");
		Map<String, String> resultMap = new HashMap<String, String>();
		String convertUpdate = updateSql.substring(updateSql.lastIndexOf("SET") + 3).replace(ExcelConstants.SQL_TAIL, "").trim();
		// logger.debug("---去掉update语句头后---" + convertUpdate);
		// 去掉WHERE子句
		convertUpdate = convertUpdate.substring(0, convertUpdate.lastIndexOf("WHERE")).trim();
		// logger.debug("---去掉where子句后---" + convertUpdate);
		String[] fieldsValues = convertUpdate.split(",");
		for (String fv : fieldsValues) {
			String[] updatedFieldValue = fv.split("=");
			resultMap.put(updatedFieldValue[0], updatedFieldValue[1]);
		}
		// 增加固定写法 wczt = 1
		resultMap.put("WCZT", "1");
		// logger.debug("---生成的map---" + resultMap);
		return resultMap;
	}

	/**
	 * 根据业务预先检验
	 *
	 * @param sysConfig
	 *            系统配置
	 * @param type
	 * @param resultMap
	 *            存放检验结果
	 * @return
	 */
	public void preCheck(Properties sysConfig, String type, Map<String, String> resultMap, Workbook workbook)
			throws Exception {
		logger.debug("---进行---预先规则检验---");
		Integer dataRowStartNum = Integer.valueOf(sysConfig.getProperty(type + "DataRowStartNum"));
		String ruleString = "";// 规则字符串
		logger.debug("---规则检验类型---" + type);
		String rulesProp = type + "CheckRules";
		ruleString = sysConfig.getProperty(rulesProp);
		checkRule(ruleString, resultMap, type, dataRowStartNum, workbook);
		logger.debug("---结束----预先规则检验---");
	}

	/**
	 * 规则检查
	 *
	 * @param ruleString
	 *            类似于：11-混合燃料:(22|23)&(24|25|26);11-双燃料:(22|23) 表示两类规则用;隔开
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
		}
		if (logger.isDebugEnabled()) {
			if (ExcelConstants.FAIL.equals(resultMap.get(ExcelConstants.RESULT_KEY))) {
				String failMsg = resultMap.get(ExcelConstants.MSG_KEY);
				logger.debug("检查规则失败：" + failMsg);
			}
		}

	}

	/**
	 * 对每一行应用检查规则
	 *
	 * @param resultMap
	 * @param checkRules
	 * @param sheet
	 * @param rowIndex
	 */
	private static void checkRowRules(Map<String, String> resultMap, String[] checkRules, Sheet sheet, int rowIndex) {
		logger.debug("---预先检验的行号---"+rowIndex);
		Row row = sheet.getRow(rowIndex);
		if (row == null) {
			logger.debug("+++行："+(rowIndex+1)+"---为空+++");
			return;
		}else {
			Cell cphmCell = row.getCell(0);//获得每一行的第一列即车牌号码
			if (StringUtils.isBlank(ExcelImportUtil.getCellValue(cphmCell))) {
				logger.debug("+++行："+(rowIndex+1)+"---的车牌号码列为空，不做规则检查跳过");
				return;
			}
		}

		for (int i = 0; i < checkRules.length; i++) {// 验证规则
			String[] ruleCell = checkRules[i].split(":");
			String[] typeCell = ruleCell[0].split("-");// ruleCell[0]为类型部分，例如：11-混合燃料
			Integer colIndex = Integer.valueOf(typeCell[0]) - 1;// typeCell[0]为类型部分的列号
			Cell tCell = row.getCell(colIndex);// 类型部分单元格
			String cellValue = ExcelImportUtil.getCellValue(tCell);
			if (StringUtils.isNotBlank(cellValue)) {

				if (cellValue.equals(typeCell[1])) {// 若单元格的值与配置的类型值（typeCell[1]）相等则进行条件校验
					if (ruleCell[1].contains("@")) {// @后面的列号代表该列不能有值
						String nonIncludeCols = ruleCell[1].split("@")[1];

						if (nonIncludeCols.contains("(")) {// 去掉括号
							nonIncludeCols = nonIncludeCols.replace("(", "").replace(")", "");
						}
						checkNonIncludeCols(resultMap, rowIndex, row, nonIncludeCols);

						String ruleCheckCols = ruleCell[1].split("@")[0];
						ruleCell[1] = ruleCheckCols;
					}


					if (ruleCell[1].contains("&")) {
						// ruleCell[1]为条件部分，其中包涵“并且”结构
						String[] conditions = ruleCell[1].split("&");

						for (int j = 0; j < conditions.length; j++) {
							if (conditions[j].contains("(")) {// 若包涵()
																// 说明()中包涵多项“或”结构
								String condition = conditions[j].replace("(", "").replace(")", "");// 去掉括号
								checkOr(resultMap, row, condition);
								if (ExcelConstants.FAIL.equals(resultMap.get(ExcelConstants.RESULT_KEY))) {
									break;
								}
							} else {// 若不包涵括号则为单值
								String nullCellNum = "";// 单元格为空的列号用逗号分割
								String colNum = conditions[j];
								Cell cCell = row.getCell(Integer.valueOf(colNum) - 1);
								if (!ExcelImportUtil.checkCellType(cCell, resultMap)) {
									continue;
								}
								if (StringUtils.isBlank(ExcelImportUtil.getCellValue(cCell))) {
									nullCellNum = ExcelImportUtil.colNumToColName(Integer.valueOf(colNum)) ;
									ExcelImportUtil.setFailMsg(resultMap, "根据燃料类型判断，第" + (rowIndex + 1) + "行，第"
											+ nullCellNum + "列，必须有值");
									break;
								}
							}
						}
					} else {// 不包涵“并且”结构 那么只有 “或”结构或者单项
						if (ruleCell[1].contains("|")) {
							checkOr(resultMap, row, ruleCell[1]);
						} else {// 不包括“或”结构 只有单项
							Integer cellIndex = Integer.valueOf(ruleCell[1]) - 1;
							Cell cCell = row.getCell(Integer.valueOf(ruleCell[1]) - 1);
							if (cCell == null || StringUtils.isBlank(ExcelImportUtil.getCellValue(cCell))) {
								String failMsg = ExcelImportUtil.makeFailMsg(rowIndex, cellIndex, "必须有值");
								ExcelImportUtil.setFailMsg(resultMap, "根据燃料类型判断，" + failMsg);
							}
						}
					}
					break;
				}
			} else {
				String failMsg = ExcelImportUtil.makeFailMsg(rowIndex, colIndex, "必须有值");
				ExcelImportUtil.setFailMsg(resultMap, failMsg);
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
				// logger.debug("---方法---checkNonIncludeCols---不包括的列号---" +
				// nonInc[j]);

				Cell cCell = row.getCell(Integer.valueOf(nonInc[j]) - 1);
				if (!ExcelImportUtil.checkCellType(cCell, resultMap)) {
					continue;
				}
				String cellValue = ExcelImportUtil.getCellValue(cCell);
				boolean isNull = true;
				if (StringUtils.isNotBlank(cellValue)) {
					isNull = false;
				}
				if (!isNull) {
					nonCellNum += ExcelImportUtil.colNumToColName(Integer.valueOf(nonInc[j])) + ",";
					isFail = true;
				}
			}
			if (isFail) {
				if (nonCellNum.contains(",")) {
					nonCellNum = nonCellNum.substring(0, nonCellNum.lastIndexOf(","));
				}

				failMsg.append("根据油料类型判断，第" + (rowIndex + 1) + "行，第" + nonCellNum + "列中不能有值");
				ExcelImportUtil.setFailMsg(resultMap, failMsg.toString());
				logger.warn(resultMap.get(ExcelConstants.MSG_KEY));
			}

		} else {
			Cell cCell = row.getCell(Integer.valueOf(nonIncludeCols) - 1);
			if (cCell != null) {
				ExcelImportUtil
						.setFailMsg(resultMap, "根据油料类型判断，第" + (rowIndex + 1) + "行，第" + nonIncludeCols + "列中不能有值");
				logger.warn(resultMap.get(ExcelConstants.MSG_KEY));
			}
		}
	}

	/**
	 * 验证规则中的“或”结构的验证
	 *
	 * @param resultMap
	 * @param rowIndex
	 * @param row
	 * @param nullCellNum
	 * @param condition
	 */
	private static void checkOr(Map<String, String> resultMap, Row row, String condition) {
		String nullCellNum = "";// 单元格为空的列号用逗号分割
		String[] colNum = condition.split("\\|");// 将多项“或”结构分割为列号
		String notNullCellNum = "";// 单元格不为空的列号用逗号分割
		int orCount = 0;// “或”结构中不为空的单元格的个数
		for (int k = 0; k < colNum.length; k++) {
			Cell cCell = row.getCell(Integer.valueOf(colNum[k]) - 1);
			if (!ExcelImportUtil.checkCellType(cCell, resultMap)) {
				continue;
			}
			if (cCell != null && StringUtils.isNotBlank(ExcelImportUtil.getCellValue(cCell))) {
				orCount++;
				notNullCellNum += ExcelImportUtil.colNumToColName(Integer.valueOf(colNum[k])) + "或";
			} else {
				nullCellNum += ExcelImportUtil.colNumToColName(Integer.valueOf(colNum[k])) + "或";
			}
		}
		if (orCount > 1) {// 或结构中只能有1列的有值
			notNullCellNum = StringUtils.substring(notNullCellNum, 0, StringUtils.lastIndexOf(notNullCellNum, "或"));
			ExcelImportUtil.setFailMsg(resultMap, "根据燃料类型判断，第" + (row.getRowNum() + 1) + "行，第" + notNullCellNum
					+ "列中只能有一列有值");
			logger.warn(resultMap.get(ExcelConstants.MSG_KEY));
		} else if (orCount == 0 && StringUtils.isNotBlank(nullCellNum)) {
			nullCellNum = StringUtils.substring(nullCellNum, 0, StringUtils.lastIndexOf(nullCellNum, "或"));
			ExcelImportUtil.setFailMsg(resultMap, "根据燃料类型判断，第" + (row.getRowNum() + 1) + "行，第" + nullCellNum
					+ "列中必须有一列有值");
			logger.warn(resultMap.get(ExcelConstants.MSG_KEY));
		}
	}




	/**
	 * 计算各燃料的百公里平均单耗
	 *
	 * @param updateMap
	 *            计算时取得的各字段的值来源，以及计算完毕后各燃料的百公里平均单耗保存的map
	 */
	private static void calBGLPJDH(Map<String, String> updateMap) {
		Double dyxslc = Double.valueOf(updateMap.get("DYXSLC"));
		calBGLPJDH(dyxslc, updateMap);
	}

	/**
	 * 根据燃料类型计算各燃料的百公里平均单耗
	 *
	 * @param dyxslc
	 *            当月行驶里程
	 * @param updateMap
	 *            计算时取得的各字段的值来源，以及计算完毕后各燃料的百公里平均单耗保存的map
	 */
	private static void calBGLPJDH(Double dyxslc, Map<String, String> updateMap) {
		// logger.debug("---开始计算百公里平均单耗---");
		String rllx = updateMap.get("RLLX");// 取得燃料类型
		// logger.debug("---燃料类型---" + rllx);
		if ("'汽油'".equals(rllx)) {
			calBGLPJDHNormal("DYQYXHL", "QYBGLPJDH", dyxslc, updateMap);
		} else if ("'柴油'".equals(rllx)) {
			calBGLPJDHNormal("DYCYXHL", "CYBGLPJDH", dyxslc, updateMap);
		} else if ("'LNG'".equals(rllx)) {
			calBGLPJDHNormal("DYLNGXHL", "LNGBGLPJDH", dyxslc, updateMap);
		} else if ("'LPG'".equals(rllx)) {
			calBGLPJDHNormal("DYLPGXHL", "LPGBGLPJDH", dyxslc, updateMap);
		} else if ("'CNG'".equals(rllx)) {
			calBGLPJDHNormal("DYCNGXHL", "CNGBGLPJDH", dyxslc, updateMap);
		} else if (StringUtils.contains(rllx, "双燃料")) {
			// 从当月汽油消耗量（DYQYXHL）或当月柴油消耗量（DYCYXHL）中取一个非空值放入updateMap中 再从
			// LNG，LPG，CNG的当月柴油消耗量中取一个非空值放入updateMap中
			calBGLPJDHNormal("DYQYXHL", "QYBGLPJDH", dyxslc, updateMap);
			calBGLPJDHNormal("DYCYXHL", "CYBGLPJDH", dyxslc, updateMap);
			calBGLPJDHNormal("DYLNGXHL", "LNGBGLPJDH", dyxslc, updateMap);
			calBGLPJDHNormal("DYLPGXHL", "LPGBGLPJDH", dyxslc, updateMap);
			calBGLPJDHNormal("DYCNGXHL", "CNGBGLPJDH", dyxslc, updateMap);
		} else if ("混合燃料".equals(rllx)) {
			// 从当月汽油消耗量（DYQYXHL）或当月柴油消耗量（DYCYXHL）中取一个非空值放入updateMap中
			calBGLPJDHNormal("DYQYXHL", "QYBGLPJDH", dyxslc, updateMap);
			calBGLPJDHNormal("DYCYXHL", "CYBGLPJDH", dyxslc, updateMap);
		}
		// logger.debug("---结束计算百公里平均单耗---");
	}

	/**
	 * 正常燃料类型（汽油，柴油，LNG，LPG,CNG）计算百公里平均单耗
	 *
	 * @param dyxhl
	 *            读取当月消耗量字段名称
	 * @param bglpjdh
	 *            存储百公里平均单耗字段名称
	 * @param updateMap
	 *            计算完时的数据来源与计算完毕后的数据存储
	 */
	private static void calBGLPJDHNormal(String dyxhl, String bglpjdh, Double dyxslc, Map<String, String> updateMap) {
		if (updateMap.get(dyxhl) != null && StringUtils.isNotBlank(updateMap.get(dyxhl))) {
			BigDecimal dyxhlBD = new BigDecimal(updateMap.get(dyxhl));// 当月消耗量
			BigDecimal dyxslcBD = BigDecimal.valueOf(dyxslc);
			if (dyxslcBD.doubleValue() != 0) {
				// 百公里平均单耗 = 当月消耗量/当月行驶历程 * 100
				BigDecimal bglpjdhBD = dyxhlBD.divide(dyxslcBD, 4, BigDecimal.ROUND_HALF_UP).multiply(
						BigDecimal.valueOf(100));
				updateMap.put(bglpjdh, String.valueOf(bglpjdhBD.doubleValue()));// 百公里平均单耗
			}else {
				updateMap.put(bglpjdh, "0");// 百公里平均单耗
			}

		}
	}


}