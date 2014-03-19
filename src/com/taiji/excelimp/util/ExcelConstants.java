package com.taiji.excelimp.util;

import java.io.Serializable;
/**
 * excel�ĵ��빤�ߵĳ�������
 * @author zhangxin
 *
 */
public class ExcelConstants implements Serializable {

	private static final long serialVersionUID = -2093316439392651132L;
	public static final String EXCEL_FILE_TYPE = "xls";
	public static final String EXCEL2007_FILE_TYPE = "xlsx";
	// ����Ϊ���map�д�ŵ�keyֵ
	public static final String RESULT_KEY = "result";
	public static final String MSG_KEY = "msg";
	public static final String SQLS_KEY = "sqls";
	public static final String FILENAME_KEY = "fileName";

	// ����Ϊ�����ļ��еı�ǩ��������������
	public static final String ELEMENT_TEMPLATE = "template";
	public static final String ELEMENT_SHEET = "sheet";
	public static final String ELEMENT_COL = "col";
	public static final String ELEMENT_VALIDATION = "validation";
	
	//����Ϊ�����ļ��еĸ���ǩ����������
	public static final String TEMPLATE_ATTR_ID = "id";
	public static final String TEMPLATE_ATTR_WORKBOOKPROTECT = "workbookProtect";
	public static final String TEMPLATE_ATTR_CERTIFICATE = "certificate";
	
	//sheet��ǩ����
	public static final String SHEET_ATTR_NAME = "name";
	public static final String SHEET_ATTR_TABLENAME = "tableName";
	public static final String SHEET_ATTR_DATASTARTROWNUM = "dataStartRowNum";
	public static final String SHEET_ATTR_ISVALIDATE = "isValidate";
	public static final String SHEET_ATTR_OPERATION = "operation";
	public static final String SHEET_ATTR_WHOLEROW = "wholeRow";
	public static final String SHEET_ATTR_COLCOUNT = "colCount";
	
	//sheet��������
	public static final String OPERATION_TYPE_INSERT = "INSERT";
	public static final String OPERATION_TYPE_UPDATE= "UPDATE";
	public static final String OPERATION_TYPE_MULTINSERT = "MULTINSERT";
	
	//col��ǩ����
	public static final String COL_ATTR_COLNUM = "colNum";
	public static final String COL_ATTR_TABFIELD = "tabField";
	public static final String COL_ATTR_FIELDTYPE = "fieldType";
	public static final String COL_ATTR_CONVERTER = "converter";
	public static final String COL_ATTR_ISCONDITION= "isCondition";
	public static final String COL_ATTR_REGEXPCHECKER = "regExpChecker";
	public static final String COL_ATTR_UNIQUE = "unique";
	public static final String COL_ATTR_MAXLENGTH = "maxLength";
	public static final String COL_ATTR_IGNORE = "ignore";
	
	//validation��ǩ����
	public static final String VALIDATION_ATTR_VALTYPE = "valType";
	

	// ����Ϊ����ʱ���еĵ��ֶ�����
	public static final String FIELD_TYPE_STRING = "STRING";
	public static final String FIELD_TYPE_NUM = "NUM";
	public static final String FIELD_TYPE_DATE = "DATE";

	// ����Ϊ����ʱ����֤����
	public static final String VALIDATION_NOTNULL = "NOTNULL";// �ǿ���֤
	public static final String VALIDATION_ISNUM = "NUM";// ������֤
	public static final String VALIDATION_IS_POSITIVE_NUM = "POSITIVENUM";// ������������֤
	public static final String VALIDATION_EMAIL = "EMAIL";// �����ʼ���֤
	public static final String VALIDATION_IDCARD = "IDCARD";// ���֤������֤
	//����Ϊ����֤���͵�������ʽ
	public static final String REGEXP_ISNUM = "^(-?\\d+)(\\.\\d+)?";
	public static final String REGEXP_IS_POSITIVE_NUM = "^[1-9]\\d*|^[1-9]\\d*\\.\\d*|0\\.\\d*[1-9]\\d*$";//��������
	
	public static final String REGEXP_DATE = "^(\\d{4})-(0\\d{1}|1[0-2])-(0\\d{1}|[12]\\d{1}|3[01])$";//���ڸ�ʽУ�� ��ʽΪ��yyyy-MM-dd
	public static final String DATE_FORMAT = "yyyy-MM-dd";
	
	//����Ϊ��֤ʧ��ԭ��
	public static final String VAL_FAIL_MSG_ISNUM = "ֵ��Ϊ����"; 
	public static final String VAL_FAIL_MSG_ISPOSITIVENUM = "ֵ��Ϊ����"; 
	public static final String VAL_FAIL_MSG_EMAIL = "email��֤ʧ��";
	public static final String VAL_FAIL_MSG_IDCARD="���֤��֤ʧ��";
	public static final String VAL_FAIL_MSG_NOTNULL="ֵΪ��";		
	
	
	
	public static final String SUCCESS = "success";// �ɹ�
	public static final String FAIL = "fail";// ʧ��
	public static final String SQL_TAIL = ";\n";//sql����β�ַ�
	public static final String SQL_INSERT_VALUE_FLAG=") @%&#VALUES#&%@ (";//���insert����value��ʶ
	public static final String SQL_INSERT_VALUE = ") VALUES (";//insert ����е�values��ʶ
	public static final String SQL_MULTINSERT_FROM_DUAL_FLAG = "@%&FROM@%&DUAL@%&";//multinsert��������������ɵı�ʶ ��������������sqlʱ�滻
	public static final String SQL_MULTINSERT_UNION_SELECT_FLAG = "@%&UNION@%&SELECT@%&";//multinsert��������������ɵı�ʶ ��������������sqlʱ
	public static final String SQL_MULTINSERT_FROM_DUAL = " FROM DUAL ";//һ��insert���������¼ʱ�̶�д�� from dual
	public static final String SQL_MULTINSERT_UNION_SELECT = " UNION SELECT ";//һ��insert���������¼ʱ�̶�д�� union select
}
