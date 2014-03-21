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
 * 车辆信息批量注册导入功能
 *
 * @author zhangxin
 *
 */
public class RegisterBatchImpExcel extends AbstractImpExcel {

	public static Logger logger = LoggerFactory.getLogger(RegisterBatchImpExcel.class);

	@Override
	public void importExcel(Properties sysConfig, DBAccess dbAccess) throws Exception {
		logger.debug("---车辆信息批量导入开始---");
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
			// 车辆信息导入的Excel文件命名规范：UUID_操作类别_单位Id_省代码_市代码_县代码_单位名称.xlsx(.xls)
			// 操作类别包括：csgjplzc,nckyplzc,czqcplzc
			// 也对应着配置文件中的template元素中的templateId属性
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

				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					// 根据车牌号码查重
					isSuccess = !checkDuplicate(insertSqls, resultMap, dbAccess, hylb);
				}

				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY)) && ("csgj".equalsIgnoreCase(hylb)||"ncky".equalsIgnoreCase(hylb))) {
					// 根据线路名称查询线路表中的线路是否存在
					insertSqls = xlmcCheck(insertSqls, resultMap, dbAccess, hylb);
					logger.debug("---线路检查完成后重组sql---"+insertSqls);
				}

				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					// 若生成insert语句成功则执行插入操作
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
						logger.debug("---最终生成的执行sql---");
						for (int j = 0; j < inserts.length; j++) {
							logger.debug(inserts[j]);
						}
					}
					// 进行批量插入
					dbAccess.batchExecuteSqls(inserts);
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
	 * 检查线路名称
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
					xlmc = valueParts[13];// 获得城市公交线路名称
					tableName = "T_CSGJ_XLGL";
				} else if ("ncky".equalsIgnoreCase(hylb)) {
					xlmc = valueParts[13];// 获得农村客运线路名称
					tableName = "T_NCKY_XLGL";
				}
				logger.debug("---在insert语句中获得的线路名称---"+xlmc);

				String bxId = getBXIDByXlmc(xlmc, conn, dbAccess, tableName);
				if (StringUtils.isNotBlank(bxId)) {// 若线路表中的记录存在则将线路id加入到insert语句中
					statePart += "," + "YYXLH";
					valuePart += "," + bxId;
					inserts[i] = insertPrefix + statePart + ExcelConstants.SQL_INSERT_VALUE_FLAG + valuePart + ")";
					logger.debug("---线路检查完毕---线路表中存在记录：" + bxId + "---重组的insert语句---" + inserts[i]);
					result.append(inserts[i]);
					result.append(ExcelConstants.SQL_TAIL);
				} else {
					String failMsg = "名称为：" + xlmc + "的线路不存在";
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
	 * 根据线路名称获得线路id
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
	 * 根据车牌号码 车牌颜色 变更情况 查询车辆信息是否有重复
	 *
	 * @param insertSqls
	 *            解析excel生成的insert语句
	 * @param resultMap
	 *            结果集合
	 * @param dbAccess
	 *            数据库访问对象
	 */
	private boolean checkDuplicate(String insertSqls, Map<String, String> resultMap, DBAccess dbAccess, String type) throws Exception{
		logger.debug("---查询重复车牌号码---");
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
			throw new Exception("无法确定查询重复车牌的表名称");
		}
		try {
			for (int i = 0; i < inserts.length; i++) {
				logger.debug("---insert语句---" + inserts[i]);
				String valuePart = inserts[i].substring(inserts[i].lastIndexOf(ExcelConstants.SQL_INSERT_VALUE_FLAG))
						.replace(ExcelConstants.SQL_INSERT_VALUE_FLAG, "");
				String [] valueParts = valuePart.split(",");
				// 获得cphm字段对应的值
				String cphmVal = valueParts[0].replace("'", "");
				//车牌颜色
				String cpysVal = valueParts[1].replace("'", "");
				//变更情况
				String bgqkVal = valueParts[2].replace("'", "");
				Map<String, String> valueMap = new HashMap<String, String>();
				valueMap.put("CPHM", cphmVal);
				valueMap.put("CPYS", cpysVal);
				valueMap.put("BGQK", bgqkVal);
				logger.debug("---车牌号码---" + cphmVal+"---车牌颜色---"+cpysVal+"---变更情况---"+bgqkVal);
				if (dbAccess.isFieldValueDup(tableName, valueMap, conn)) {
					ExcelImportUtil.setFailMsg(resultMap, "车牌号码为" + cphmVal + "，车牌颜色为"+cpysVal+"，变更情况为"+bgqkVal+"的记录已经存在无法导入");
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
		//计算车龄
		valuePart = calCL(valuePart);
		statePart += "," + "CL";

		String[] valueParts = valuePart.split(",");

		//根据燃料类型取舍燃料类型相关字段
		result = calRllx(statePart, valuePart, insertPreFix, valueParts);
		logger.debug("---预先重组生成的insert语句---" + result);
		return result;
	}

	/**
	 * 计算车龄
	 * @param valuePart
	 * @return
	 */
	private String calCL(String valuePart) {
		// 计算车龄
		String fzDateVal = valuePart.split(",")[6].replace("'", "");// 获得发证日期的值
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
			logger.debug("---车龄计算结果---"+cl);
		} catch (ParseException e) {
			logger.error(e.getMessage(), e);
		}
		return valuePart;
	}

	/**
	 * 燃料类型计算
	 * @param statePart
	 * @param valuePart
	 * @param insertPreFix
	 * @param valueParts
	 * @return
	 */
	private String calRllx(String statePart, String valuePart, String insertPreFix, String[] valueParts) {
		String result;
		// 根据然辆类型 算出 rllx1或者rllx2或者rllx3字段
		String rllx = "";
		if (StringUtils.containsIgnoreCase(insertPreFix, "CSGJJCB")) {
			rllx = valueParts[23];
		}else if (StringUtils.containsIgnoreCase(insertPreFix, "NCKYJCB")) {
			rllx = valueParts[23];
		}else if (StringUtils.containsIgnoreCase(insertPreFix, "CZQCJCB")) {
			rllx = valueParts[23];
		}

		if ("'汽油'".equals(rllx)) {
			statePart = StringUtils.replace(statePart, "RLLX*1,", "");
			valuePart = StringUtils.replace(valuePart, ",'单燃料-汽油'", "");
		} else if ("'柴油'".equals(rllx)) {
			statePart = StringUtils.replace(statePart, "RLLX*1,", "");
			valuePart = StringUtils.replace(valuePart, ",'单燃料-柴油'", "");
		} else if ("'LPG'".equals(rllx)) {
			statePart = StringUtils.replace(statePart, "RLLX*1,", "");
			valuePart = StringUtils.replace(valuePart, ",'单燃料-LPG'", "");
		} else if ("'CNG'".equals(rllx)) {
			statePart = StringUtils.replace(statePart, "RLLX*1,", "");
			valuePart = StringUtils.replace(valuePart, ",'单燃料-CNG'", "");
		} else if ("'LNG'".equals(rllx)) {
			statePart = StringUtils.replace(statePart, "RLLX*1,", "");
			valuePart = StringUtils.replace(valuePart, ",'单燃料-LNG'", "");
		} else if ("'双燃料'".equals(rllx)) {
			statePart = StringUtils.replace(statePart, "RLLX*1", "RLLX1" + "," + "RLLX2");
			if (StringUtils.contains(valuePart, "'汽油+LPG'")) {
				valuePart = StringUtils.replace(valuePart, "'汽油+LPG'", "'汽油','LPG'");
			} else if (StringUtils.contains(valuePart, "'汽油+CNG'")) {
				valuePart = StringUtils.replace(valuePart, "'汽油+CNG'", "'汽油','CNG'");
			} else if (StringUtils.contains(valuePart, "'汽油+LNG'")) {
				valuePart = StringUtils.replace(valuePart, "'汽油+LNG'", "'汽油','LNG'");
			} else if (StringUtils.contains(valuePart, "'柴油+LPG'")) {
				valuePart = StringUtils.replace(valuePart, "'柴油+LPG'", "'柴油','LPG'");
			} else if (StringUtils.contains(valuePart, "'柴油+CNG'")) {
				valuePart = StringUtils.replace(valuePart, "'柴油+CNG'", "'柴油','CNG'");
			} else if (StringUtils.contains(valuePart, "'柴油+LNG'")) {
				valuePart = StringUtils.replace(valuePart, "'柴油+LNG'", "'柴油','LNG'");
			}
		} else if ("'混合动力'".equals(rllx)) {
			statePart = StringUtils.replace(statePart, "RLLX*1", "RLLX3");
		}

		// 重新组合
		result = insertPreFix + statePart + ExcelConstants.SQL_INSERT_VALUE_FLAG + valuePart + ")";
		return result;
	}

}