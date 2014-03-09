package com.taiji.excelimp.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.taiji.excelimp.ImpCheck;
import com.taiji.excelimp.db.DBAccess;

/**
 * 行政区划工具类
 * 
 * @author zhangxin
 * 
 */
public class RegionUtil {
	private static Logger logger = LoggerFactory.getLogger(RegionUtil.class);
	private static HashMap<String, Long> shiMap;
	private static HashMap<String, Long> xianMap;
	
	/**
	 * 初始化shiMap 和 xianMap
	 * @param sMap
	 * @param xMap
	 */
	public static void initShiXianMap(HashMap<String, Long>sMap, HashMap<String, Long>xMap){
		shiMap = sMap;
		xianMap = xMap;
	}

	/**
	 * 获得地市缓存Map
	 */
	public static HashMap<String, Long> getShiMap(){
		return shiMap;
	}
	/**
	 * 获得县缓存Map
	 */
	public static HashMap<String, Long>getXianMap(){
		return xianMap;
	}
	
	/**
	 * 该方法用于定期更新农村客运模板中的县市行政区划的数据源
	 * 
	 * @param excelDataSourceFile
	 * @param dbAccess
	 */
	public static void genRegionExcelDataSource(File excelDataSourceFile, DBAccess dbAccess) {
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet();
		Row shiRow = sheet.createRow(0);
		Row shiCodeRow = sheet.createRow(1);
		Connection conn = null;
		Statement stmt = null;
		Statement stmtXian = null;
		ResultSet rs = null;
		ResultSet rsXian = null;
		try {
			FileOutputStream fos = new FileOutputStream(excelDataSourceFile);
			String shiSql = "select distinct(shi),shiname from t_sys_user_count t where t.sheng='650000' and shiname is not null  order by shi";
			conn = dbAccess.getConnection();
			stmt = conn.createStatement();
			rs = stmt.executeQuery(shiSql);
			for (int i = 0; rs.next(); i++) {
				Cell cell = shiRow.createCell(i);
				cell.setCellValue(new XSSFRichTextString(rs.getString(2)));
				
				String xianSql = "select xian,xianname from t_sys_user_count t where t.shi=" + rs.getLong(1)
						+ " and xian <> -1 order by xian";
				stmtXian = conn.createStatement();
				rsXian = stmtXian.executeQuery(xianSql);
				int xianCount = 0;
				for (int j = 0; rsXian.next(); j++) {
					xianCount++;
					Row xianRow = null;
					if (sheet.getRow(j + 2) == null) {
						xianRow = sheet.createRow(j + 2);
					} else {
						xianRow = sheet.getRow(j + 2);
					}
					Cell xianCell = xianRow.createCell(i);
					if (0==j) {
						xianCell.setCellValue(new XSSFRichTextString("市直"));
					}else {
						xianCell.setCellValue(new XSSFRichTextString(rsXian.getString(2)));
					}
				}
				//地市名称下方为地市代码-地市下的县数目：如：652700_3
				cell = shiCodeRow.createCell(i);
				cell.setCellValue(new XSSFRichTextString(rs.getString(1)+"_"+xianCount));
				
			}
			workbook.write(fos);
		} catch (SQLException e) {
			logger.error(e.getMessage(), e);
		} catch (IOException e) {
			logger.error(e.getMessage(), e);
		} finally {
			dbAccess.release(rsXian, stmtXian);
			dbAccess.release(rs, stmt);
			dbAccess.release(conn);
		}

	}

	public static void main(String[] args) {
		try {
			String propPath = RegionUtil.class.getResource("/").getPath() + "impsysconfig.properties";
			Properties config = ImpCheck.readProperties(propPath);
			DBAccess dbAccess = new DBAccess(config.getProperty("dburl"), config.getProperty("username"), config.getProperty("password"),
					config.getProperty("driverClassName"));
			File expPath = new File(config.getProperty("datasourceFilePath"));
			if (!expPath.exists()) {
				expPath.mkdirs();
			}
			File expFile = new File(expPath,"datasource.xlsx");
			
			RegionUtil.genRegionExcelDataSource(expFile, dbAccess);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
