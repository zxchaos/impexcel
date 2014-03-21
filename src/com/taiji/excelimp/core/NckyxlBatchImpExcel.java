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
 * 农村客运线路批量导入业务类
 * @author zhangxin
 *
 */
public class NckyxlBatchImpExcel extends AbstractImpExcel {
	@Override
	public void importExcel(Properties sysConfig, DBAccess dbAccess) throws Exception {
		logger.debug("---农村客运线路批量导入开始---");
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
			// 车辆信息导入的Excel文件命名规范：UUID_操作类别_单位Id_地市代码.xlsx(.xls)
			// 操作类别包括：nckyxlgl
			// 也对应着配置文件中的template元素中的templateId属性
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
				
				//获得要导入的文件的工作簿对象并检查文件的有效性
				workbook = ExcelImportUtil.genWorkbook(impFiles[i], document, templateId, resultMap);
				
				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					//开始解析文件
					ExcelImportUtil.importExcel(workbook, document, templateId, resultMap);
				}
				
				insertSqls = resultMap.get(ExcelConstants.SQLS_KEY);
				if (!ExcelConstants.FAIL.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY)) && StringUtils.isBlank(insertSqls)) {
					logger.info("+++生成的sql为空+++可能是模板中没有数据");
					ExcelImportUtil.setFailMsg(resultMap, "导入的模板中不包含数据");
				}
				
				if (ExcelConstants.SUCCESS.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
					//若生成insert语句成功则执行插入操作
					Map<String, Object>fieldValueMap = new HashMap<String, Object>();
					fieldValueMap.put("XS", Double.valueOf(0.1));
					fieldValueMap.put("ZT", Long.valueOf(0));
					fieldValueMap.put("SHENG", Long.valueOf(650000));
					fieldValueMap.put("SHI", Long.valueOf(fileNameParts[3]));//从文件名称中取得地市代码
					String [] inserts = super.remakeInsert(insertSqls, fieldValueMap,"BXID",dbAccess);
					//进行批量插入
					dbAccess.batchExecuteSqls(inserts);
					super.insertImpInfo(dbAccess, resultMap, infoFieldMap,true,super.getType());
					isSuccess = true;
				}else{
					//若生成失败将错误信息写入数据库
					isSuccess = false;
					super.insertImpInfo(dbAccess, resultMap, infoFieldMap, isSuccess,super.getType());
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
			}finally{
				//将处理完成的文件移动到备份目录
				backupFile(impFiles[i], backupDir, isSuccess);
			}
		}
	}

}