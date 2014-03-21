package com.taiji.excelimp.util;

import java.io.Serializable;
/**
 * excel的导入工具的常量配置
 * @author zhangxin
 *
 */
public class ExcelConstants implements Serializable {

	private static final long serialVersionUID = -2093316439392651132L;
	public static final String EXCEL_FILE_TYPE = "xls";
	public static final String EXCEL2007_FILE_TYPE = "xlsx";
	// 以下为结果map中存放的key值
	public static final String RESULT_KEY = "result";
	public static final String MSG_KEY = "msg";
	public static final String SQLS_KEY = "sqls";
	public static final String FILENAME_KEY = "fileName";

	// 以下为配置文件中的标签名称与属性名称
	public static final String ELEMENT_TEMPLATE = "template";
	public static final String ELEMENT_SHEET = "sheet";
	public static final String ELEMENT_COL = "col";
	public static final String ELEMENT_VALIDATION = "validation";
	
	//以下为配置文件中的各标签的属性名称
	public static final String TEMPLATE_ATTR_ID = "id";
	public static final String TEMPLATE_ATTR_WORKBOOKPROTECT = "workbookProtect";
	public static final String TEMPLATE_ATTR_CERTIFICATE = "certificate";
	
	//sheet标签属性
	public static final String SHEET_ATTR_NAME = "name";
	public static final String SHEET_ATTR_TABLENAME = "tableName";
	public static final String SHEET_ATTR_DATASTARTROWNUM = "dataStartRowNum";
	public static final String SHEET_ATTR_ISVALIDATE = "isValidate";
	public static final String SHEET_ATTR_OPERATION = "operation";
	public static final String SHEET_ATTR_WHOLEROW = "wholeRow";
	public static final String SHEET_ATTR_COLCOUNT = "colCount";
	
	//sheet操作类型
	public static final String OPERATION_TYPE_INSERT = "INSERT";
	public static final String OPERATION_TYPE_UPDATE= "UPDATE";
	public static final String OPERATION_TYPE_MULTINSERT = "MULTINSERT";
	
	//col标签属性
	public static final String COL_ATTR_COLNUM = "colNum";
	public static final String COL_ATTR_TABFIELD = "tabField";
	public static final String COL_ATTR_FIELDTYPE = "fieldType";
	public static final String COL_ATTR_CONVERTER = "converter";
	public static final String COL_ATTR_ISCONDITION= "isCondition";
	public static final String COL_ATTR_REGEXPCHECKER = "regExpChecker";
	public static final String COL_ATTR_UNIQUE = "unique";
	public static final String COL_ATTR_MAXLENGTH = "maxLength";
	public static final String COL_ATTR_IGNORE = "ignore";
	
	//validation标签属性
	public static final String VALIDATION_ATTR_VALTYPE = "valType";
	

	// 以下为导入时表中的的字段类型
	public static final String FIELD_TYPE_STRING = "STRING";
	public static final String FIELD_TYPE_NUM = "NUM";
	public static final String FIELD_TYPE_DATE = "DATE";

	// 以下为导入时的验证类型
	public static final String VALIDATION_NOTNULL = "NOTNULL";// 非空验证
	public static final String VALIDATION_ISNUM = "NUM";// 数字验证
	public static final String VALIDATION_IS_POSITIVE_NUM = "POSITIVENUM";// 正浮点数字验证
	public static final String VALIDATION_EMAIL = "EMAIL";// 电子邮件验证
	public static final String VALIDATION_IDCARD = "IDCARD";// 身份证号码验证
	//以下为各验证类型的正则表达式
	public static final String REGEXP_ISNUM = "^(-?\\d+)(\\.\\d+)?";
	public static final String REGEXP_IS_POSITIVE_NUM = "^[1-9]\\d*|^[1-9]\\d*\\.\\d*|0\\.\\d*[1-9]\\d*$";//正浮点数
	
	public static final String REGEXP_DATE = "^(\\d{4})-(0\\d{1}|1[0-2])-(0\\d{1}|[12]\\d{1}|3[01])$";//日期格式校验 格式为：yyyy-MM-dd
	public static final String DATE_FORMAT = "yyyy-MM-dd";
	
	//以下为验证失败原因
	public static final String VAL_FAIL_MSG_ISNUM = "值不为数字"; 
	public static final String VAL_FAIL_MSG_ISPOSITIVENUM = "值不为正数"; 
	public static final String VAL_FAIL_MSG_EMAIL = "email验证失败";
	public static final String VAL_FAIL_MSG_IDCARD="身份证验证失败";
	public static final String VAL_FAIL_MSG_NOTNULL="值为空";		
	
	
	
	public static final String SUCCESS = "success";// 成功
	public static final String FAIL = "fail";// 失败
	public static final String SQL_TAIL = ";\n";//sql语句结尾字符
	public static final String SQL_INSERT_VALUE_FLAG=") @%&#VALUES#&%@ (";//拆分insert语句的value标识
	public static final String SQL_INSERT_VALUE = ") VALUES (";//insert 语句中的values标识
	public static final String SQL_MULTINSERT_FROM_DUAL_FLAG = "@%&FROM@%&DUAL@%&";//multinsert操作解析完后生成的标识 用于生成真正的sql时替换
	public static final String SQL_MULTINSERT_UNION_SELECT_FLAG = "@%&UNION@%&SELECT@%&";//multinsert操作解析完后生成的标识 用于生成真正的sql时
	public static final String SQL_MULTINSERT_FROM_DUAL = " FROM DUAL ";//一条insert插入多条记录时固定写法 from dual
	public static final String SQL_MULTINSERT_UNION_SELECT = " UNION SELECT ";//一条insert插入多条记录时固定写法 union select
}