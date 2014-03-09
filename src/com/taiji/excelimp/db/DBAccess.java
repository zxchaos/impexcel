package com.taiji.excelimp.db;

import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;

import oracle.jdbc.OracleTypes;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class DBAccess {
	public static Logger logger = LoggerFactory.getLogger(DBAccess.class);
	private String dbUrl;
	private String userName;
	private String password;
	private String driverClassName;

	public DBAccess(String dbUrl, String userName, String password, String driverClassName) {

		this.dbUrl = dbUrl;
		this.userName = userName;
		this.password = password;
		this.driverClassName = driverClassName;
	}

	public Connection getConnection() {
		Connection conn = null;
		try {
			long start = System.currentTimeMillis();
			Class.forName(this.driverClassName);
			conn = DriverManager.getConnection(dbUrl, userName, password);
			long end = System.currentTimeMillis();
			logger.debug("----获得conn---经历时间" + (end - start));
		} catch (ClassNotFoundException e) {
			logger.error(e.getMessage(), e);
		} catch (SQLException e) {
			logger.error(e.getMessage(), e);
		}
		return conn;
	}

	/**
	 * 批量执行sql
	 * 
	 * @param sqls
	 */
	public void batchExecuteSqls(String[] sqls) throws Exception {
		int result = 0;
		Connection conn = null;
		Statement stmt = null;
		try {
			conn = this.getConnection();
			conn.setAutoCommit(false);
			stmt = conn.createStatement();
			for (int i = 0; i < sqls.length; i++) {
				stmt.addBatch(sqls[i]);
			}
			int [] count = stmt.executeBatch();
			conn.commit();
		} catch (SQLException e) {
			logger.error(e.getMessage(), e);
			rollback(conn);
			throw e;
		} catch (Exception e) {
			logger.error(e.getMessage(), e);
			rollback(conn);
			throw e;
		} finally {
			release(stmt, conn);
		}

	}
	
	/**
	 * 批量执行sql
	 * 
	 * @param sqls
	 */
	public void batchExecuteSqls(String[] sqls, Connection conn) throws Exception {
		int result = 0;
		Statement stmt = null;
		try {
			stmt = conn.createStatement();
			for (int i = 0; i < sqls.length; i++) {
				stmt.addBatch(sqls[i]);
			}
			int [] count = stmt.executeBatch();
			conn.commit();
		} catch (SQLException e) {
			logger.error(e.getMessage(), e);
			rollback(conn);
			throw e;
		} catch (Exception e) {
			logger.error(e.getMessage(), e);
			rollback(conn);
			throw e;
		} finally {
			release(stmt);
		}

	}

	/**
	 * 执行sql
	 * 
	 * @param sql
	 * @return
	 * @throws Exception
	 */
	public int executeSql(String sql) throws Exception {
		int result = 0;
		Connection conn = null;
		Statement stmt = null;
		try {
			conn = this.getConnection();
			conn.setAutoCommit(false);
			stmt = conn.createStatement();
			result = stmt.executeUpdate(sql);
			conn.commit();
		} catch (SQLException e) {
			logger.error(e.getMessage(), e);
			rollback(conn);
			throw e;
		} catch (Exception e) {
			logger.error(e.getMessage(), e);
			rollback(conn);
			throw e;
		} finally {
			release(stmt, conn);
		}
		return result;
	}

	/**
	 * @param v_jb
	 *            当前用户级别
	 * @param v_xzjb
	 *            行政级别
	 * @param v_proviceid
	 *            省ID
	 * @param v_cityid
	 *            市ID
	 * @param v_countyid
	 *            县ID
	 * @param v_year
	 *            年度
	 * @param v_month
	 *            月份
	 * @param v_province
	 *            省名称
	 * @param v_city
	 *            市名称
	 * @param v_county
	 *            县名称
	 * @param v_company
	 *            公司名称
	 * 
	 * */
	public String updateSjhz(int v_jb, int v_xzjb, long v_proviceid, long v_cityid, long v_countyid,
			String v_companyid, int v_year, int v_month, String v_province, String v_city, String v_county,
			String v_company, Connection conn) throws Exception {

		if (conn == null) {
			return "0,no conn";
		}

		int v_hzjb = v_jb;
		if (v_jb == 9 && v_xzjb == 2) {
			// 市直
			v_hzjb = 3;
			v_county = "市直";
			v_countyid = -1;
		}
		if (v_jb == 9 && v_xzjb == 1) {
			// 省直
			v_hzjb = 2;
			v_city = "省直";
			v_cityid = -1;
			v_county = "省直";
			v_countyid = -1;
		}
		int v_dwlb = v_jb;

		int v_out_int = -1;
		String v_out_str = "";
		CallableStatement proc = null;
		try {
			proc = conn.prepareCall("{call p_sjhz_csgj_gen_manual_new(?,?,?,?,?,?,?,?,?,?,?,?,?,?)}");
			proc.setInt(1, v_xzjb);
			proc.setInt(2, v_dwlb);
			proc.setLong(3, v_proviceid);
			proc.setLong(4, v_cityid);
			proc.setLong(5, v_countyid);
			proc.setString(6, v_companyid);
			proc.setInt(7, v_year);
			proc.setInt(8, v_month);
			proc.setString(9, v_province);
			proc.setString(10, v_city);
			proc.setString(11, v_county);
			proc.setString(12, v_company);
			proc.registerOutParameter(13, OracleTypes.INTEGER);
			proc.registerOutParameter(14, OracleTypes.VARCHAR);

			proc.execute();

			v_out_int = proc.getInt(13);
			v_out_str = proc.getString(14);

			// 调试信息
			if (v_out_str != null)
				logger.debug(v_out_str);
			return v_out_int + "," + v_out_str;
		} catch (Exception ex) {
			ex.printStackTrace();
			conn.rollback();
			throw ex;
		} finally {
			proc.close();
		}
	}
	
	/**
	 * 执行批量更新和调用存储过程
	 * @param dbAccess
	 * @param updateSqls
	 * @param fileNameParts
	 * @throws Exception
	 */
	public void  updateAndCallprocedure(String [] updateSqls, String [] fileNameParts) throws Exception{
		logger.debug("---开始批量更新和调用存储过程---");
		Connection conn = this.getConnection();
		try {
			conn.setAutoCommit(false);
			// 将解析excel生成的sql导入数据库
			logger.debug("---将update语句写入数据库---");
			this.batchExecuteSqls(updateSqls,conn);
			logger.debug("---写入数据库完毕---");

			logger.debug("---调用存储过程---");
			this.updateSjhz(9, 4, Long.valueOf(fileNameParts[2]), Long.valueOf(fileNameParts[3]),
					Long.valueOf(fileNameParts[4]), fileNameParts[5], Integer.parseInt(fileNameParts[6]),
					Integer.parseInt(fileNameParts[7]), fileNameParts[8], fileNameParts[9], fileNameParts[10],
					fileNameParts[11],conn);
			logger.debug("---结束调用存储过程---");
			conn.commit();
		} catch (Exception e) {
			conn.rollback();
			logger.debug("---批量更新和调用存储过程出错---");
			logger.error(e.getMessage(),e);
			throw e;
		}finally{
			this.release(conn);
		}
		logger.debug("---结束批量更新和调用存储过程---");
	}
	

	/**
	 * 释放Statement Connection资源
	 * 
	 * @param stmt
	 *            要被释放的Statement
	 * @param conn
	 *            要被释放的Connection
	 */
	public void release(Statement stmt, Connection conn) {
		release(stmt);
		release(conn);
	}

	/**
	 * 释放conn
	 * 
	 * @param conn
	 */
	public  void release(Connection conn) {
		if (conn != null) {
			try {
				conn.close();
				conn = null;
			} catch (SQLException e) {
				logger.error(e.getMessage(), e);
			}
		}
	}

	/**
	 * 释放Statement
	 * 
	 * @param stmt
	 *            要被释放的Statement
	 */
	public  void release(Statement stmt) {
		if (stmt != null) {
			try {
				stmt.close();
				stmt = null;
			} catch (SQLException e) {
				logger.error(e.getMessage(), e);
			}
		}
	}

	/**
	 * 释放资源
	 * 
	 * @param rs
	 *            要被释放的ResultSet
	 */
	public void release(ResultSet rs) {
		if (rs != null) {
			try {
				rs.close();
				rs = null;
			} catch (SQLException e) {
				logger.error("ResultSet release failed", e);
			}
		}
	}

	/**
	 * 释放资源
	 * 
	 * @param rs
	 *            要被释放的ResultSet
	 * @param stmt
	 *            要被释放的Statement
	 */
	public void release(ResultSet rs, Statement stmt) {
		release(rs);
		release(stmt);
	}

	public static void rollback(Connection conn) {
		try {
			conn.rollback();
		} catch (SQLException e) {
			logger.error(e.getMessage(), e);
		}
	}

	public String getOneFieldContent(String sql) {
		Connection conn = null;
		String tmp = null;
		try {
			conn = getConnection();
			tmp = getOneFieldContent(sql, conn);
		} catch (Exception e) {
			logger.error("运行SQL=" + sql + "时出错", e);
		} finally {
			release(conn);
		}
		return org.apache.commons.lang3.StringUtils.trim(tmp);
	}

	/**
	 * 取得数据库一条记录，适用于sql语句中只查询一个字段
	 * 
	 * @param sql
	 *            SQL语句
	 * @param conn
	 *            数据库连接
	 * @return List
	 */
	public String getOneFieldContent(String sql, Connection conn) throws Exception{
		String tmp = null;
		Statement stmt = null;
		ResultSet rs = null;
		try {
			stmt = conn.createStatement();
			rs = stmt.executeQuery(sql);
			// ArrayList al = new ArrayList();
			while (rs.next()) {
				tmp = rs.getString(1);
			}
			rs.close();
			rs = null;
			stmt.close();
			stmt = null;

		} catch (Exception ex) {
			logger.error("error:", ex);
			logger.debug("error sql=" + sql);
			throw ex;
		} finally {
			release(rs, stmt);
		}
		return org.apache.commons.lang3.StringUtils.trim(tmp);
	}

	public synchronized long getSequence() throws Exception {
		long seq = Long.valueOf(getNowDateString());
		long dbseq = Long.valueOf(getOneFieldContent(" select SEQ_ID.nextval from dual "));
		return seq * 100000 + dbseq;
	}

	public synchronized long getSequence(Connection conn) throws Exception {
		long seq = Long.valueOf(getNowDateString());
		long dbseq = Long.valueOf(getOneFieldContent(" select SEQ_ID.nextval from dual ", conn));
		return seq * 100000 + dbseq;
	}
	
	/**
	 * 获得当前时间yyyyMMdd
	 * 
	 * @return 返回当前时间yyyyMMdd格式
	 */
	private String getNowDateString() {
		SimpleDateFormat sdf = new SimpleDateFormat("MMddHHmm");
		return sdf.format(new Date());
	}

	/**
	 * 将解析出错信息插入数据库
	 * 
	 * @param dwid
	 * @param hylb
	 * @param sj
	 * @param info
	 */
	public void insertErrorInfo(long dwid, String hylb, String sj, String info,String type) throws Exception {
		String insertsql = "INSERT INTO PLDRFK_INFO (ID,DWID,HYLB,SJ,INFO,TYPE) values(?,?,?,?,?,?)";
		logger.debug("---插入导入结果信息的预编译sql---"+insertsql);
		Connection conn = null;
		PreparedStatement psmt = null;
		try {
			long id = getSequence();
			logger.debug("---生成id---"+id);
			conn = getConnection();
			psmt = conn.prepareStatement(insertsql);
			psmt.setLong(1, id);
			psmt.setLong(2, dwid);
			psmt.setString(3, hylb);
			psmt.setString(4, sj);
			psmt.setString(5, info);
			psmt.setString(6, type);
			psmt.executeUpdate();
			logger.debug("---执行插入完成---");
		} catch (SQLException e) {
			logger.error("执行sql出错，"+e.getMessage(),e);
			throw e;
		} catch (Exception e) {
			logger.error(e.getMessage(),e);
			throw e;
		}finally{
			release(psmt,conn);
		}
	}
	
	/**
	 * 判断某字段的值是否在某表中有重复
	 * @param fieldName
	 * @param value
	 * @return
	 */
	public boolean isFieldValueDup(String tableName,String fieldName, String value, Connection conn ) throws Exception{
		boolean result = true;
		PreparedStatement psmt = null;
		ResultSet rs = null;
		String sql = "SELECT SJID FROM "+tableName+" WHERE "+fieldName+"=?";
		logger.debug("---查重sql---"+sql);
		try {
			psmt = conn.prepareStatement(sql);
			psmt.setString(1, value);
			rs = psmt.executeQuery();
			if (rs.next()) {
				result = true;
				logger.debug("---有"+fieldName+"为"+value+"的重复的记录---");
			}else {
				result = false;
				logger.debug("---没有"+fieldName+"为"+value+"的记录---");
			}
		} catch (SQLException e) {
			logger.error(e.getMessage(),e);
			throw e;
		}finally{
			release(rs);
			release(psmt);
		}
		return result;
	}
	
	
	
	/**
	 * @param strSQL
	 * @return
	 */
	public HashMap<String,Long> initHashMap(String strSQL){
		logger.debug("---开始获取缓存Map---");
		logger.debug("---获取sql---"+strSQL);
		Connection conn = null;
		PreparedStatement pstm = null;
		ResultSet rst = null;
		
		HashMap<String,Long> hashmp = new HashMap<String,Long>();
		try {
			conn = getConnection();
			pstm = conn.prepareStatement(strSQL);
			rst =  pstm.executeQuery();
			
			while(rst.next()){
				Long shiOrXian = rst.getLong(1);
				String shiOrXiaNname = rst.getString(2);
				hashmp.put(shiOrXiaNname,shiOrXian);
			}
		} catch (Exception e) {
			logger.error(e.getMessage(),e);
		} finally {
			this.release(conn);
		}
		logger.debug("---获取map完毕---");
		return hashmp;
	}
	/**
	 * 插入受益人表 和 基础表
	 * @param conn
	 * @param multInsertSql
	 * @param pks
	 * @param sjid
	 */
	public void multInsertSyrAndJCB(Connection conn, String multInsertSql,String pks, String sjid, String hylb) throws Exception{
		Statement stmtSyr = null;
		Statement stmtJcb = null;
		try {
			stmtSyr = conn.createStatement();
			stmtSyr.executeUpdate(multInsertSql);
			logger.debug("---multinsert---插入受益人表完毕---");
			String jcbSql = "UPDATE "+hylb+"JCB SET ";
			String [] pkParts = pks.split(",");
			for (int i = 0; i < pkParts.length; i++) {
				jcbSql += "BTSYRID"+(i+1)+"="+pkParts[i]+",";
			}
			jcbSql = jcbSql.substring(0, jcbSql.lastIndexOf(","));
			jcbSql += " WHERE SJID="+sjid;
			logger.debug("---multinsert---插入基础表sql---"+jcbSql);
			stmtJcb = conn.createStatement();
			stmtJcb.executeUpdate(jcbSql);
			logger.debug("---multinsert---插入基础表完毕---");
			
		} catch (SQLException e) {
			logger.error(e.getMessage(),e);
			throw e;
		}finally{
			release(stmtSyr);
			release(stmtJcb);
		}
	}
	
	public static void main(String [] args){
		DBAccess dbAccess = new DBAccess("jdbc:oracle:thin:@192.168.10.17:1521:orcl", "rybt_xj", "rybt_xj", "oracle.jdbc.driver.OracleDriver");
//		System.out.println(dbAccess.isFieldValueDup("CSGJJCB", "DWID", "1123225005222"));
	}
}
