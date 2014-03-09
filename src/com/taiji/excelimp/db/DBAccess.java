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
			logger.debug("----���conn---����ʱ��" + (end - start));
		} catch (ClassNotFoundException e) {
			logger.error(e.getMessage(), e);
		} catch (SQLException e) {
			logger.error(e.getMessage(), e);
		}
		return conn;
	}

	/**
	 * ����ִ��sql
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
	 * ����ִ��sql
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
	 * ִ��sql
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
	 *            ��ǰ�û�����
	 * @param v_xzjb
	 *            ��������
	 * @param v_proviceid
	 *            ʡID
	 * @param v_cityid
	 *            ��ID
	 * @param v_countyid
	 *            ��ID
	 * @param v_year
	 *            ���
	 * @param v_month
	 *            �·�
	 * @param v_province
	 *            ʡ����
	 * @param v_city
	 *            ������
	 * @param v_county
	 *            ������
	 * @param v_company
	 *            ��˾����
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
			// ��ֱ
			v_hzjb = 3;
			v_county = "��ֱ";
			v_countyid = -1;
		}
		if (v_jb == 9 && v_xzjb == 1) {
			// ʡֱ
			v_hzjb = 2;
			v_city = "ʡֱ";
			v_cityid = -1;
			v_county = "ʡֱ";
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

			// ������Ϣ
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
	 * ִ���������º͵��ô洢����
	 * @param dbAccess
	 * @param updateSqls
	 * @param fileNameParts
	 * @throws Exception
	 */
	public void  updateAndCallprocedure(String [] updateSqls, String [] fileNameParts) throws Exception{
		logger.debug("---��ʼ�������º͵��ô洢����---");
		Connection conn = this.getConnection();
		try {
			conn.setAutoCommit(false);
			// ������excel���ɵ�sql�������ݿ�
			logger.debug("---��update���д�����ݿ�---");
			this.batchExecuteSqls(updateSqls,conn);
			logger.debug("---д�����ݿ����---");

			logger.debug("---���ô洢����---");
			this.updateSjhz(9, 4, Long.valueOf(fileNameParts[2]), Long.valueOf(fileNameParts[3]),
					Long.valueOf(fileNameParts[4]), fileNameParts[5], Integer.parseInt(fileNameParts[6]),
					Integer.parseInt(fileNameParts[7]), fileNameParts[8], fileNameParts[9], fileNameParts[10],
					fileNameParts[11],conn);
			logger.debug("---�������ô洢����---");
			conn.commit();
		} catch (Exception e) {
			conn.rollback();
			logger.debug("---�������º͵��ô洢���̳���---");
			logger.error(e.getMessage(),e);
			throw e;
		}finally{
			this.release(conn);
		}
		logger.debug("---�����������º͵��ô洢����---");
	}
	

	/**
	 * �ͷ�Statement Connection��Դ
	 * 
	 * @param stmt
	 *            Ҫ���ͷŵ�Statement
	 * @param conn
	 *            Ҫ���ͷŵ�Connection
	 */
	public void release(Statement stmt, Connection conn) {
		release(stmt);
		release(conn);
	}

	/**
	 * �ͷ�conn
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
	 * �ͷ�Statement
	 * 
	 * @param stmt
	 *            Ҫ���ͷŵ�Statement
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
	 * �ͷ���Դ
	 * 
	 * @param rs
	 *            Ҫ���ͷŵ�ResultSet
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
	 * �ͷ���Դ
	 * 
	 * @param rs
	 *            Ҫ���ͷŵ�ResultSet
	 * @param stmt
	 *            Ҫ���ͷŵ�Statement
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
			logger.error("����SQL=" + sql + "ʱ����", e);
		} finally {
			release(conn);
		}
		return org.apache.commons.lang3.StringUtils.trim(tmp);
	}

	/**
	 * ȡ�����ݿ�һ����¼��������sql�����ֻ��ѯһ���ֶ�
	 * 
	 * @param sql
	 *            SQL���
	 * @param conn
	 *            ���ݿ�����
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
	 * ��õ�ǰʱ��yyyyMMdd
	 * 
	 * @return ���ص�ǰʱ��yyyyMMdd��ʽ
	 */
	private String getNowDateString() {
		SimpleDateFormat sdf = new SimpleDateFormat("MMddHHmm");
		return sdf.format(new Date());
	}

	/**
	 * ������������Ϣ�������ݿ�
	 * 
	 * @param dwid
	 * @param hylb
	 * @param sj
	 * @param info
	 */
	public void insertErrorInfo(long dwid, String hylb, String sj, String info,String type) throws Exception {
		String insertsql = "INSERT INTO PLDRFK_INFO (ID,DWID,HYLB,SJ,INFO,TYPE) values(?,?,?,?,?,?)";
		logger.debug("---���뵼������Ϣ��Ԥ����sql---"+insertsql);
		Connection conn = null;
		PreparedStatement psmt = null;
		try {
			long id = getSequence();
			logger.debug("---����id---"+id);
			conn = getConnection();
			psmt = conn.prepareStatement(insertsql);
			psmt.setLong(1, id);
			psmt.setLong(2, dwid);
			psmt.setString(3, hylb);
			psmt.setString(4, sj);
			psmt.setString(5, info);
			psmt.setString(6, type);
			psmt.executeUpdate();
			logger.debug("---ִ�в������---");
		} catch (SQLException e) {
			logger.error("ִ��sql����"+e.getMessage(),e);
			throw e;
		} catch (Exception e) {
			logger.error(e.getMessage(),e);
			throw e;
		}finally{
			release(psmt,conn);
		}
	}
	
	/**
	 * �ж�ĳ�ֶε�ֵ�Ƿ���ĳ�������ظ�
	 * @param fieldName
	 * @param value
	 * @return
	 */
	public boolean isFieldValueDup(String tableName,String fieldName, String value, Connection conn ) throws Exception{
		boolean result = true;
		PreparedStatement psmt = null;
		ResultSet rs = null;
		String sql = "SELECT SJID FROM "+tableName+" WHERE "+fieldName+"=?";
		logger.debug("---����sql---"+sql);
		try {
			psmt = conn.prepareStatement(sql);
			psmt.setString(1, value);
			rs = psmt.executeQuery();
			if (rs.next()) {
				result = true;
				logger.debug("---��"+fieldName+"Ϊ"+value+"���ظ��ļ�¼---");
			}else {
				result = false;
				logger.debug("---û��"+fieldName+"Ϊ"+value+"�ļ�¼---");
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
		logger.debug("---��ʼ��ȡ����Map---");
		logger.debug("---��ȡsql---"+strSQL);
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
		logger.debug("---��ȡmap���---");
		return hashmp;
	}
	/**
	 * ���������˱� �� ������
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
			logger.debug("---multinsert---���������˱����---");
			String jcbSql = "UPDATE "+hylb+"JCB SET ";
			String [] pkParts = pks.split(",");
			for (int i = 0; i < pkParts.length; i++) {
				jcbSql += "BTSYRID"+(i+1)+"="+pkParts[i]+",";
			}
			jcbSql = jcbSql.substring(0, jcbSql.lastIndexOf(","));
			jcbSql += " WHERE SJID="+sjid;
			logger.debug("---multinsert---���������sql---"+jcbSql);
			stmtJcb = conn.createStatement();
			stmtJcb.executeUpdate(jcbSql);
			logger.debug("---multinsert---������������---");
			
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
