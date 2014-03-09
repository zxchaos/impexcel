package com.taiji.excelimp.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FilenameFilter;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.Node;
import org.dom4j.io.SAXReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.taiji.excelimp.api.IConverter;
import com.taiji.excelimp.api.IRegExpChecker;

/**
 * Excel 导入工具类 本类完成对所导入的Excel文件的工作簿结合导入配置文件import.xml中的配置进行导入与验证
 * 
 * @author zhangxin
 * 
 */
public class ExcelImportUtil {
	public static final Logger logger = LoggerFactory.getLogger(ExcelImportUtil.class);

	/**
	 * 导入Excel功能入口方法 将目录impExcelDir下的所有的excel文件按照配置文件中的同一个模板导入
	 * 
	 * @param impExcelDir
	 *            存放Excel文件的目录
	 * @param templateId
	 *            要导入的excel文件对应的配置文件中template标签中的id属性值
	 * @param configFilePath
	 *            配置文件路径
	 * @param excelFileType
	 *            目录下的excel文件类型 分为 xls 和 xlsx 两种
	 * 
	 * @return 返回中的Map的key包含"result":表示解析成功或失败其值为"fail"或"success" 若解析失败即result对应的值为fail时 则包含key："msg" 其中包含错误信息
	 *         若解析成功即result对应的值为success时 则包含key："sqls" 其中包含解析成功后的insert 或 update 语句 insert 或 update 语句格式为: 
	 *         INSERT INTO TABLENAME (FIELD1,FIELD2...) @%&#VALUES#&%@ (VALUE1,VALUE2);\n 
	 *         或者 UPDATE TABLENAME SET FIELD1=VALUE1,FIELD2=VALUE2... WHERE FIELD3=VALUE3 AND FIELD4=VALUE4 ...;\n
	 */
	public static Map<String, String> importExcel(String impExcelDir, String templateId, String configFilePath,
			String excelFileType) throws Exception {
		File excelDir = validateFile(impExcelDir, true);
		File[] excelFiles = getDirFiles(excelDir, excelFileType);
		Map<String, String> resultMap = new HashMap<String, String>();// 存放解析结果的Map

		for (File excelFile : excelFiles) {
			Map<String, String> impMap = importExcel(excelFile, templateId, configFilePath);
			if (ExcelConstants.FAIL.equalsIgnoreCase(impMap.get(ExcelConstants.RESULT_KEY))) {
				appendFailMsg(resultMap, excelFile, impMap);
				continue;
			} else if (ExcelConstants.SUCCESS.equalsIgnoreCase(impMap.get(ExcelConstants.RESULT_KEY))) {
				setSuccessMsg(resultMap, impMap);
			}
		}
		return resultMap;
	}

	/**
	 * 导入Excel功能入口方法 根据配置文件中模板id和配置文件路径导入Excel
	 * 
	 * @param excelFile
	 *            要导入的excel文件
	 * @param templateId
	 *            模板Id
	 * @param configFilePath
	 *            配置文件位置
	 * @return 返回中的Map的key包含"result":表示解析成功或失败其值为"fail"或"success" 若解析失败即result对应的值为fail时 则包含key："msg" 其中包含错误信息
	 *         若解析成功即result对应的值为success时 则包含key："sqls" 其中包含解析成功后的insert语句 insert 语句格式为: 
	 *         INSERT INTO TABLENAME (FIELD1,FIELD2...) @%&#VALUES#&%@ (VALUE1,VALUE2);\n 
	 *         或者 
	 *         UPDATE TABLENAME SET FIELD1=VALUE1,FIELD2=VALUE2...;\n
	 */
	public static Map<String, String> importExcel(File excelFile, String templateId, String configFilePath)
			throws Exception {
		logger.debug("+++开始---方法---importExcel");
		Map<String, String> resultMap = new HashMap<String, String>();
		Document document = getConfigFileDoc(configFilePath);
		Workbook workbook = genWorkbook(excelFile, document, templateId, resultMap);
		importExcel(workbook, document, templateId);
		logger.debug("+++结束---方法---importExcel");
		return resultMap;
	}
	
	
	public static void importExcel(Workbook workbook,Document configDoc, String templateId)
			throws Exception {
		importExcel(workbook, configDoc, templateId, null);
	}
	
	/**
	 * 导入Excel功能入口方法 根据配置文件中模板id和配置文件路径导入Excel
	 * 
	 * @param workbook
	 *            要解析的工作簿对象
	 * @param configDoc 导入配置文件的文档对象
	 * @param templateId
	 *            模板Id
	 * @param resultMap 存放解析结果
	 * @return 返回中的Map的key包含"result":表示解析成功或失败其值为"fail"或"success" 若解析失败即result对应的值为fail时 则包含key："msg" 其中包含错误信息
	 *         若解析成功即result对应的值为success时 则包含key："sqls" 其中包含解析成功后的insert语句 insert 语句格式为: 
	 *         INSERT INTO TABLENAME (FIELD1,FIELD2...) @%&#VALUES#&%@ (VALUE1,VALUE2);\n 
	 *         或者 
	 *         UPDATE TABLENAME SET FIELD1=VALUE1,FIELD2=VALUE2...;\n
	 *         或者
	 *         INSERT INTO TABLENAME (FIELD1,FIELD2...) SELECT VALUE1,VALUE2... FROM DUAL UNION SELECT ...
	 */
	public static Map<String, String> importExcel(Workbook workbook,Document configDoc, String templateId, Map<String, String> resultMap)
			throws Exception {
		logger.debug("+++开始---方法---importExcel");
		if (resultMap == null) {
			resultMap = new HashMap<String, String>();
		}
		doImport(workbook, configDoc, templateId,resultMap);
		logger.debug("+++结束---方法---importExcel");
		return resultMap;
	}

	/**
	 * 设置excel导入成功信息
	 * 
	 * @param resultMap
	 *            总结果map
	 * @param impMap
	 *            某次导入后存放的导入信息map
	 */
	public static void setSuccessMsg(Map<String, String> resultMap, Map<String, String> impMap) {
		if (resultMap.get(ExcelConstants.SQLS_KEY) != null) {
			StringBuffer sqls = new StringBuffer(resultMap.get(ExcelConstants.SQLS_KEY));
			sqls.append(impMap.get(ExcelConstants.SQLS_KEY));
			resultMap.put(ExcelConstants.SQLS_KEY, sqls.toString());
		} else {
			resultMap.put(ExcelConstants.SQLS_KEY, impMap.get(ExcelConstants.SQLS_KEY));
		}
		resultMap.put(ExcelConstants.RESULT_KEY, ExcelConstants.SUCCESS);
	}

	/**
	 * 设置resultMap并追加失败信息的文件信息
	 * 
	 * @param resultMap
	 *            结果map
	 * @param excelFile
	 *            导入的文件信息
	 * @param impMap
	 *            文件导入结果
	 */
	public static void appendFailMsg(Map<String, String> resultMap, File excelFile, Map<String, String> impMap) {
		if (resultMap.get(ExcelConstants.MSG_KEY) != null) {
			StringBuffer msg = new StringBuffer("");
			msg.append(resultMap.get(ExcelConstants.MSG_KEY));
			msg.append(impMap.get(ExcelConstants.MSG_KEY));
			resultMap.put(ExcelConstants.MSG_KEY, msg.toString());
		}else {
			resultMap.put(ExcelConstants.MSG_KEY, impMap.get(ExcelConstants.MSG_KEY));
		}
		resultMap.put(ExcelConstants.RESULT_KEY, ExcelConstants.FAIL);
		if (null != excelFile ) {
			resultMap.put(ExcelConstants.FILENAME_KEY, excelFile.getName());// 保存文件名
		}
		resultMap.remove(ExcelConstants.SQLS_KEY);
	}

	/**
	 * 将resultMap的失败信息追加到finalResultMap中
	 * @param finalResultMap
	 * @param resultMap
	 */
	public static void appendFailMsg(Map<String, String> finalResultMap, Map<String, String> resultMap){
		appendFailMsg(finalResultMap, null, resultMap);
	}
	/**
	 * 获得目录下的指定扩展名的文件
	 * 
	 * @param excelDir
	 *            要获得指定类型文件的目录
	 * @param fileExtName
	 *            要获得的文件类型的扩展名
	 * @return
	 */
	public static File[] getDirFiles(File excelDir, final String fileExtName) {
		FilenameFilter filenameFilter = new FilenameFilter() {
			public boolean accept(File dir, String name) {
				Pattern pattern = Pattern.compile("[\\s\\S]*.(" + fileExtName + ")$");
				return pattern.matcher(name).matches();
			}
		};

		File[] excelFiles = excelDir.listFiles(filenameFilter);
		return excelFiles;
	}

	/**
	 * 获得目录下的的excel文件包括Excel 2007+与2003
	 * 
	 * @param excelDir
	 *            要获得excel文件的目录
	 * @return
	 */
	public static File[] getDirFiles(File excelDir) {
		FilenameFilter filenameFilter = new FilenameFilter() {
			public boolean accept(File dir, String name) {
				boolean result = false;
				Pattern pattern2003 = Pattern.compile("[\\s\\S]*.(" + ExcelConstants.EXCEL_FILE_TYPE + ")$");
				Pattern pattern2007 = Pattern.compile("[\\s\\S]*.(" + ExcelConstants.EXCEL2007_FILE_TYPE + ")$");
				result = pattern2003.matcher(name).matches();
				if (!result) {// 若文件不是2003 接着判断是否为2007
					result = pattern2007.matcher(name).matches();
				}
				return result;
			}
		};

		File[] excelFiles = excelDir.listFiles(filenameFilter);
		return excelFiles;
	}

	/**
	 * 获得导入配置文件的文档对象
	 * 
	 * @param configFilePath
	 * @return 配置文件的文档对象
	 * @throws Exception
	 * @throws DocumentException
	 */
	public static Document getConfigFileDoc(String configFilePath) throws Exception, DocumentException {
		long start = System.currentTimeMillis();
		File configFile = validateFile(configFilePath, false);
		SAXReader reader = new SAXReader();
		Document document = reader.read(configFile);
		long end = System.currentTimeMillis();
		logger.debug("---成功获得配置文件文档对象---经历时间---"+(end-start));
		return document;
	}

	/**
	 * 导入Excel 工作簿
	 * 
	 * @param workbook
	 *            工作簿对象
	 * @param doc
	 *            配置文件的xml文档对象
	 * @param templateId
	 *            配置文件中模板对应的id值
	 * @param isValidate
	 *            是否进行导入验证
	 * @return 导入结果集合(取返回map中的key为 ExcelConstants.RESULT_KEY
	 *         对应的value值{"success","fail"}可以判断导入是否成功,若result对应value为"fail"则导入失败，再获取map中的key为ExcelConstants.MSG_KEY
	 *         对应的value可获得导入出错信息)
	 */
	public static void doImport(Workbook workbook, Document doc, String templateId, Map<String, String> resultMap) throws Exception {
		logger.debug("---开始---方法---doImport");
		List<Element> sheetEleList = getSheetList(doc, templateId);

		for (Element sheetEle : sheetEleList) {
			String sheetName = sheetEle.attributeValue(ExcelConstants.SHEET_ATTR_NAME);// 获得配置文件中template标签下的sheet的名称即在配置文件中可导入的sheet名称
			if (sheetName != null && !"".equals(sheetName)) {
				Sheet impSheet = workbook.getSheet(sheetName);// 通过名称获得workbook中的将要导入的sheet
				if (impSheet != null) {
					impSheet(impSheet, sheetEle,resultMap);
					if (ExcelConstants.FAIL.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
						break;
					}
				} else {
					String failMsg = "名称：" + sheetName + "没有对应的sheet";
					setFailMsg(resultMap, failMsg);
					break;
				}
			}
		}
		logger.debug("+++结束---方法---doImport");
	}

	/**
	 * @param resultMap
	 * @param failMsg
	 */
	public static void setFailMsg(Map<String, String> resultMap, String failMsg) {
		setFailMsg(resultMap, failMsg, true);
	}
	
	/**
	 * @param resultMap
	 * @param failMsg
	 * @param isAppend 是否追加错误信息
	 */
	public static void setFailMsg(Map<String, String> resultMap, String failMsg,boolean isAppend) {
		resultMap.put(ExcelConstants.RESULT_KEY, ExcelConstants.FAIL);
		if (resultMap.get(ExcelConstants.MSG_KEY) == null) {
			resultMap.put(ExcelConstants.MSG_KEY, failMsg+"\n");
		} else {
			if (isAppend) {
				String origMsg = resultMap.get(ExcelConstants.MSG_KEY);
				resultMap.put(ExcelConstants.MSG_KEY, origMsg + failMsg + "\n");
			} else {
				String origMsg = resultMap.get(ExcelConstants.MSG_KEY);
				resultMap.put(ExcelConstants.MSG_KEY, failMsg+"\n"+origMsg + "\n");
			}
		}
	}

	/**
	 * 根据配置文件的文档对象和配置文件中template标签id值获得某一template标签下的sheet标签
	 * 
	 * @param doc
	 * @param templateId
	 * @return
	 */
	public static List<Element> getSheetList(Document doc, String templateId) {
		Element root = doc.getRootElement();
		Node templatNode = root.selectSingleNode("//template[@" + ExcelConstants.TEMPLATE_ATTR_ID + "='"
				+ templateId + "']");
		List<Element> sheetEleList = templatNode.selectNodes("./" + ExcelConstants.ELEMENT_SHEET);// 要导入的sheet
		return sheetEleList;
	}

	/**
	 * 根据配置文件sheet标签中的配置导入excel文件sheet中的数据
	 * 
	 * @param impSheet
	 * @param sheetEle
	 * @param isValidate
	 *            导入时是否进行校验
	 * @return insert语句集合
	 */
	public static Map<String, String> impSheet(Sheet impSheet, Element sheetEle,Map<String, String>resultMap) throws Exception {
		logger.debug("---开始---方法---impSheet");
		StringBuffer genSqls = new StringBuffer();// 存放组装完成的insert或update语句

		Integer dataStartRowNum = Integer.valueOf(sheetEle
				.attributeValue(ExcelConstants.SHEET_ATTR_DATASTARTROWNUM)) - 1;// 导入数据起始行数

		Boolean isValidate = Boolean.valueOf(sheetEle.attributeValue(ExcelConstants.SHEET_ATTR_ISVALIDATE));// 导入的sheet是否进行验证
		Boolean isWholeRow = Boolean.valueOf(sheetEle.attributeValue(ExcelConstants.SHEET_ATTR_WHOLEROW));//导入的列是否为一整行
		
		StringBuffer sql = null;
		String operation = sheetEle.attributeValue(ExcelConstants.SHEET_ATTR_OPERATION);
		boolean isInsert = true;
		int colCount = 0;
		if (operation.equalsIgnoreCase(ExcelConstants.OPERATION_TYPE_INSERT)) {// 配置insert操作
			sql = makePartInsertSql(sheetEle);
		} else if (operation.equalsIgnoreCase(ExcelConstants.OPERATION_TYPE_UPDATE)) {// 配置update操作
			sql = makePartUpdateSql(sheetEle);
			isInsert = false;
		} else if (ExcelConstants.OPERATION_TYPE_MULTINSERT.equalsIgnoreCase(operation)) {//multinser操作
			sql = makePartMultInsertSql(sheetEle);
			colCount = Integer.valueOf(sheetEle.attributeValue(ExcelConstants.SHEET_ATTR_COLCOUNT));
			isInsert = false;
		} else {
			throw new Exception("sheet 标签的operation属性设置错误");
		}

		List<Element> colList = sheetEle.selectNodes("./" + ExcelConstants.ELEMENT_COL);// 获得sheet元素下所有的col子元素
		List<String> uniqueList = new ArrayList<String>();
		
		for (int i = dataStartRowNum; i <= impSheet.getLastRowNum(); i++) {
			StringBuffer rowSql = new StringBuffer(sql);// 导入的sheet中每一行组成一个insert 或者 update 语句
			Map<String, String> rowResultMap = new HashMap<String, String>();
			Row dataRow = impSheet.getRow(i);
			if(dataRow == null){
				logger.warn("+++行："+(i+1)+"为空+++");
				continue;
			}
			logger.debug("---解析行---" + (dataRow.getRowNum()+1));
			
			StringBuffer updateWhere = new StringBuffer("");// update操作的条件
			int nullColNum = 0;//值为空的列数
			int multInsertCount = 0;//multinsert操作时计数
			int multInsertBlankNum = 0;//multinsert操作时的空列的数目
			for (Element colEle : colList) {
				Integer colNum = Integer.valueOf(colEle.attributeValue(ExcelConstants.COL_ATTR_COLNUM).trim()) - 1;// 获得要导入的列的列号
				
				Cell impCell = dataRow.getCell(colNum);
				
				if (!checkCellType(impCell, rowResultMap)) {
					continue;
				}
				
				//要导入的列是否为空
				if (null == impCell || StringUtils.isBlank(getCellValue(impCell))) {
					nullColNum++;
				}
				
				if (isValidate != null && isValidate 
						&& !validating(rowResultMap, colEle, impCell, impSheet, dataRow)) {// 导入的sheet需要验证
						continue;
				}
				//检查正则表达式
				if (!regExpCheck(colEle, impCell,rowResultMap,dataRow)) {
					continue;
				}
				
				//检查excel中的重复值
				if (checkUnique(rowResultMap, uniqueList, colEle, impCell)) {
					continue;
				}
				
				//检查所填值的最大长度
				if (!checkLength(rowResultMap, colEle, impCell)) {
					continue;
				}
				Boolean isCondition = Boolean
						.valueOf(colEle.attributeValue(ExcelConstants.COL_ATTR_ISCONDITION));
				if (isCondition != null && isCondition) {// 该单元格作为update条件
					makeUpdateWhere(dataRow,updateWhere, colEle, impCell, rowResultMap);
				} else if (ExcelConstants.OPERATION_TYPE_MULTINSERT.equalsIgnoreCase(operation)) {//若操作为multinsert
					multInsertCount++;
					String cellValue = getCellValue(impCell);
					if (StringUtils.isBlank(cellValue)) {
						multInsertBlankNum++;
					}
					if (multInsertCount == colCount) {
						multInsertCount = 0;
						if (multInsertBlankNum < colCount) {
							makeMultInsertSql(dataRow, rowSql, colEle, impCell, rowResultMap, true);
						}else {
							logger.debug("---multinsert---该组的值都为空---此时的insertsql为---"+rowSql);
							rowSql.replace(rowSql.lastIndexOf(ExcelConstants.SQL_MULTINSERT_UNION_SELECT_FLAG)+ExcelConstants.SQL_MULTINSERT_UNION_SELECT_FLAG.length(), rowSql.length(), "");
							logger.debug("---multinsert---去掉全部为空一组后的sql---"+rowSql);
						}
						multInsertBlankNum = 0;
					}else {
						makeMultInsertSql(dataRow, rowSql, colEle, impCell, rowResultMap, false);
					}
				} else {
					makeRowSql(dataRow, rowSql, colEle, impCell, rowResultMap, isInsert);
				}
			}
			int configColNum = colList.size();
			logger.debug("行"+(i+1)+"的空列数目为"+nullColNum+"---配置的列的数目为"+configColNum+"---配置中的wholeRow属性值为："+isWholeRow);
			
			if (null != isWholeRow && isWholeRow && nullColNum == configColNum) {//若一整行都为空
				logger.info("行："+(i+1)+"为空");
				continue;
			}else if (ExcelConstants.FAIL.equals(rowResultMap.get(ExcelConstants.RESULT_KEY))) {
				// 校验失败
					appendFailMsg(resultMap, rowResultMap);
					continue;
			}
			
			if (ExcelConstants.FAIL.equals(resultMap.get(ExcelConstants.RESULT_KEY))) {// 校验失败
				continue;
			}
			if (ExcelConstants.OPERATION_TYPE_INSERT.equals(operation)) {// insert操作
				rowSql.replace(rowSql.lastIndexOf(","), rowSql.length(), ")");
			} else if(ExcelConstants.OPERATION_TYPE_UPDATE.equals(operation)) {// update 操作
				makeUpdateRowSql(rowSql, updateWhere);
				if (!StringUtils.contains(rowSql, "WHERE")) {// 是否包含where
					String failMsg = "生成的update语句不包含where子句！";
					logger.error(failMsg);
					setFailMsg(resultMap, failMsg);
					break;
				}
			}else if (ExcelConstants.OPERATION_TYPE_MULTINSERT.equals(operation)) {//一条insert插入多条记录
				rowSql.replace(rowSql.lastIndexOf(ExcelConstants.SQL_MULTINSERT_UNION_SELECT_FLAG), rowSql.length(),"");
				logger.debug("---multinsert---整行处理完毕生成sql---"+rowSql);
			}
			
			genSqls.append(rowSql);
			genSqls.append(ExcelConstants.SQL_TAIL);
		}

		if (!ExcelConstants.FAIL.equals(resultMap.get(ExcelConstants.RESULT_KEY))) {
			resultMap.put(ExcelConstants.RESULT_KEY, ExcelConstants.SUCCESS);
			resultMap.put(ExcelConstants.SQLS_KEY, genSqls.toString());
		}
		logger.debug("+++结束---方法---impSheet");
		return resultMap;
	}

	/**
	 * 检验字符串长度是否超出最大长度
	 * @param resultMap
	 * @param colEle
	 * @param impCell
	 * @return true:验证成功；false：验证失败
	 */
	private static boolean checkLength(Map<String, String> resultMap, Element colEle, Cell impCell) {
		boolean result = true;
		String maxLength = colEle.attributeValue(ExcelConstants.COL_ATTR_MAXLENGTH);
		if (StringUtils.isNotBlank(maxLength)) {
			int mLength = Integer.valueOf(maxLength);
			String cellValue = getCellValue(impCell);
			//汉字匹配模式
			Pattern zhCharPattern = Pattern.compile("[\u4E00-\u9FA5\uFF00-\uFFEF\u3000-\u303F]");
			Matcher zhMatcher = zhCharPattern.matcher(cellValue);
			int length = cellValue.length();
			while (zhMatcher.find()) {
				//因为插入数据库时一个汉字等于两个字符所以要加1
				length++;
			}
			if (length > mLength) {
				setFailMsg(impCell, resultMap, "值的长度超出限制");
				result = false;
			}
		}
		return result;
	}
	/**
	 * 检查某列中的值在excel该列中是否有重复
	 * @param resultMap
	 * @param uniqueList
	 * @param colEle
	 * @param impCell
	 * @return true:存在重复；false：不存在重复
	 */
	private static boolean checkUnique(Map<String, String> resultMap, List<String> uniqueList, Element colEle,
			Cell impCell) {
		boolean result = false;
		String unique = colEle.attributeValue(ExcelConstants.COL_ATTR_UNIQUE);
		if (StringUtils.isNotBlank(unique) && Boolean.valueOf(unique)) {
			String cellValue = getCellValue(impCell);
			if (uniqueList.contains(cellValue)) {
				setFailMsg(impCell, resultMap, "'"+cellValue+"'在该列中已存在重复值");
				result = true;
			}else {
				uniqueList.add(cellValue);
				result = false;
			}
		}
		return result;
	}
	

	/**
	 * 检查导入的单元格中的值是否与给定的正则表达式检查器匹配
	 * @param colEle
	 * @param impCell
	 * @return 匹配结果
	 * @throws ClassNotFoundException 
	 * @throws IllegalAccessException 
	 * @throws InstantiationException 
	 */
	private static boolean regExpCheck(Element colEle, Cell impCell, Map<String, String>resultMap, Row dataRow) throws ClassNotFoundException, InstantiationException, IllegalAccessException {
		Element validation = (Element) colEle.selectSingleNode("./" + ExcelConstants.ELEMENT_VALIDATION);
		boolean hasValidation = hasVal(validation);
		
		boolean result = true;
		String recheckerAttrVal = colEle.attributeValue(ExcelConstants.COL_ATTR_REGEXPCHECKER);
		if (StringUtils.isNotBlank(recheckerAttrVal)) {//是否配置正则表达式检查器 并且 有验证限制
			IRegExpChecker regExpChecker = (IRegExpChecker)Class.forName(recheckerAttrVal).newInstance();
			if (impCell != null) {
				String cellValue = getCellValue(impCell);
				if (StringUtils.isNotBlank(cellValue)) {
					result = regExpChecker.check(cellValue);
					if (!result) {
						setFailMsg(impCell, resultMap, "所填写值的格式错误");
						logger.debug("---正则表达式检验失败---");
					}
				}else {
					if (hasValidation) {
						setFailMsg(impCell,resultMap, "该单元格不能为空");
						result = false;
					}
				}
			}else {
				if (hasValidation) {
					setFailMsg(resultMap, colEle, dataRow, "该单元格不能为空");
					result = false;
				}
			}
		}
		return result;
	}
	
	/**
	 * 生成要更新的行的 update sql语句
	 * 
	 * @param rowSql
	 * @param updateWhere
	 */
	private static void makeUpdateRowSql(StringBuffer rowSql, StringBuffer updateWhere) {
		if (StringUtils.isNotBlank(updateWhere)) {// 若where子句不为空
			// 将where子句最后的多余的AND去掉
			updateWhere.replace(updateWhere.lastIndexOf(" AND "), updateWhere.length(), "");
			String where = " WHERE " + updateWhere;
			rowSql.replace(rowSql.lastIndexOf(","), rowSql.length(), where);
		} else {
			rowSql.replace(rowSql.lastIndexOf(","), rowSql.length(), "");
		}
	}
	
	/**
	 * 生成multinsert语句
	 * @param row
	 * @param multInsertSql
	 * @param colEle
	 * @param cell
	 * @param resultMap
	 * @param isRecord 是否为完整的一条记录
	 */
	private static void makeMultInsertSql(Row row, StringBuffer multInsertSql, Element colEle, Cell cell,
			Map<String, String> resultMap, boolean isRecord) {
		String cellValue = getCellValue(cell);
		multInsertSql.append(transformIntoSqlFormat(row, colEle, cellValue, resultMap));
		multInsertSql.append(",");
		if (isRecord) {
			multInsertSql.replace(multInsertSql.lastIndexOf(","), multInsertSql.length(), "");
			multInsertSql.append(ExcelConstants.SQL_MULTINSERT_FROM_DUAL_FLAG);
			multInsertSql.append(ExcelConstants.SQL_MULTINSERT_UNION_SELECT_FLAG);
		}
	}

	/**
	 * 生成update操作的where条件子句 生成的where子句格式为 field1=value1 AND field1=value1 AND ...
	 * 
	 * @param updateWhere
	 * @param colEle
	 *            表中的列名
	 * @param whereCell
	 *            作为条件的单元格对象
	 */
	private static void makeUpdateWhere(Row row,StringBuffer updateWhere, Element colEle, Cell whereCell,
			Map<String, String> resultMap) {
		String cellValue = getCellValue(whereCell);
		if (StringUtils.isNotBlank(cellValue.trim())) {// 若cell值不为空则生成where条件
			String fieldName = colEle.attributeValue(ExcelConstants.COL_ATTR_TABFIELD);
			updateWhere.append(fieldName);
			updateWhere.append("=");
			updateWhere.append(transformIntoSqlFormat(row, colEle, cellValue, resultMap));
			updateWhere.append(" AND ");
		} else {
			String failMsg = "行："+(row.getRowNum()+1)+"中作为更新条件的单元格的值不能为空";
			logger.warn(failMsg);
			setFailMsg(resultMap, failMsg);
		}
	}

	/**
	 * 组成每一行要的insert 或 update sql
	 * 
	 * @param rowSql
	 *            已拼完成的部分insert 或 update sql
	 * @param colEle
	 *            配置文件中的col元素
	 * @param impCell
	 *            导入的excel文件中的单元格
	 */
	private static void makeRowSql(Row row, StringBuffer rowSql, Element colEle, Cell impCell, Map<String, String> resultMap,
			Boolean isInsert) throws Exception {
		
		String cellValue = getCellValueByConvertion(colEle, impCell,resultMap,row);
		makeFullSqlByType(row, rowSql, colEle, cellValue, resultMap, isInsert);
	}
	
	/**
	 * 根据转换器获得单元格的值
	 * 
	 * @param colEle
	 * @param impCell
	 * @return
	 * @throws InstantiationException
	 * @throws IllegalAccessException
	 * @throws ClassNotFoundException
	 * @throws Exception
	 */
	private static String getCellValueByConvertion(Element colEle, Cell impCell, Map<String, String>resultMap,Row row) throws InstantiationException,
			IllegalAccessException, ClassNotFoundException, Exception {
		String cellValue = getCellValue(impCell);
		String converterVal = colEle.attributeValue(ExcelConstants.COL_ATTR_CONVERTER);

		if (converterVal != null && !"".equals(converterVal)) {// 该列已配置转换器
			if (StringUtils.isBlank(cellValue)) {
				String failMsg = "该列不能为空";
				setFailMsg(resultMap, colEle, row, failMsg);
			}else {
				cellValue = convertCellVal(impCell, converterVal, resultMap);
			}
		}
		
		return cellValue;
	}

	/**
	 * @param impCell
	 * @param converterVal
	 * @return
	 * @throws InstantiationException
	 * @throws IllegalAccessException
	 * @throws ClassNotFoundException
	 */
	protected static String convertCellVal(Cell impCell, String converterVal, Map<String, String>resultMap) throws InstantiationException,
			IllegalAccessException, ClassNotFoundException, Exception {
		String cellValue = "";
		if (impCell != null) {
			IConverter converter = (IConverter) Class.forName(converterVal).newInstance();
			cellValue = converter.convert(impCell);
			if ("".equals(cellValue)) {
				setFailMsg(impCell, resultMap, "该单元格值的填写不符合规范");
			}
		}
		return cellValue;
	}

	/**
	 * 验证操作
	 * 
	 * @param resultMap
	 * @param colEle
	 * @param impCell
	 */
	private static boolean validating(Map<String, String> resultMap, Element colEle, Cell impCell, Sheet sheet, Row row) {
		Element validation = (Element) colEle.selectSingleNode("./" + ExcelConstants.ELEMENT_VALIDATION);
		boolean hasValidation = hasVal(validation);
		boolean isSuccess = true;
		if (hasValidation) {
			String valType = validation.attributeValue(ExcelConstants.VALIDATION_ATTR_VALTYPE);
			isSuccess = doValidate(impCell, valType, resultMap,colEle,row);
		}
		return isSuccess;
	}

	/**
	 * 是否配置验证元素
	 * @param validation
	 * @return
	 */
	private static boolean hasVal(Element validation) {
		boolean hasValidation = false;
		if (validation != null) {
			hasValidation = true;
		}
		return hasValidation;
	}

	/**
	 * 根据配置文件中的sheet元素来组成insert语句的：values 之前的部分
	 * 
	 * @param sheetEle
	 * @return
	 */
	private static StringBuffer makePartInsertSql(Element sheetEle) {
		List<Element> colList = sheetEle.selectNodes("./" + ExcelConstants.ELEMENT_COL);// 获得sheet元素下所有的col子元素
		String tableName = sheetEle.attributeValue(ExcelConstants.SHEET_ATTR_TABLENAME);
		StringBuffer insertSql = new StringBuffer("INSERT INTO ");
		insertSql.append(tableName);
		insertSql.append(" (");
		for (Element colEle : colList) {
			insertSql.append(colEle.attributeValue(ExcelConstants.COL_ATTR_TABFIELD));
			insertSql.append(",");
		}
		insertSql.replace(insertSql.lastIndexOf(","), insertSql.length(), ExcelConstants.SQL_INSERT_VALUE_FLAG);
		return insertSql;
	}

	/**
	 * 生成部分的update语句
	 * 
	 * @param sheetElement
	 *            配置文件中的sheet标签
	 * @return
	 */
	private static StringBuffer makePartUpdateSql(Element sheetElement) {
		StringBuffer updateSql = new StringBuffer("");
		String tableName = sheetElement.attributeValue(ExcelConstants.SHEET_ATTR_TABLENAME);
		updateSql.append("UPDATE ");
		updateSql.append(tableName);
		updateSql.append(" SET ");
		return updateSql;
	}
	/**
	 * 生成部分多行插入的insert sql 即：INSERT INTO TABLENAME (FIELD1,FIELD2...) SELECT
	 * @param sheetElement
	 * @return
	 */
	private static StringBuffer makePartMultInsertSql(Element sheetElement){
		StringBuffer result = new StringBuffer("");
		String tableName = sheetElement.attributeValue(ExcelConstants.SHEET_ATTR_TABLENAME);
		result.append("INSERT INTO ");
		result.append(tableName);
		result.append(" (");
		List colList = sheetElement.selectNodes("./"+ExcelConstants.ELEMENT_COL);
		Integer colCount = Integer.valueOf(sheetElement.attributeValue(ExcelConstants.SHEET_ATTR_COLCOUNT));
		for(int i = 0; i<colCount; i++){
			Element colEle = (Element)colList.get(i);
			String field = colEle.attributeValue(ExcelConstants.COL_ATTR_TABFIELD);
			result.append(field);
			result.append(",");
		}
		result = result.replace(result.lastIndexOf(","),result.length(), "");
		result.append(") SELECT@%&");
		logger.debug("---部分多行插入insert sql生成---"+result);
		return result;
	}

	/**
	 * 验证excel单元格
	 * 
	 * @param cell
	 *            要验证的excel单元格
	 * @param validateRule
	 *            验证规则
	 * @param colEle 列配置元素
	 * @param row 导入的行
	 * @return 验证是否通过
	 */
	public static boolean doValidate(Cell cell, String valType,Map<String, String>resultMap, Element colEle, Row row) {
		boolean isSuccess = true;
		if (valType.equalsIgnoreCase(ExcelConstants.VALIDATION_ISNUM)) {// 验证是否数字
			isSuccess = valRegExp(cell, resultMap, ExcelConstants.REGEXP_ISNUM, ExcelConstants.VAL_FAIL_MSG_ISNUM, row, colEle);
		} else if (valType.equalsIgnoreCase(ExcelConstants.VALIDATION_IS_POSITIVE_NUM)) {//正浮点数字验证
			isSuccess = valRegExp(cell, resultMap, ExcelConstants.REGEXP_IS_POSITIVE_NUM, ExcelConstants.VAL_FAIL_MSG_ISPOSITIVENUM, row, colEle);
		}else if (valType.equalsIgnoreCase(ExcelConstants.VALIDATION_NOTNULL)) {// 验证是否为空
			isSuccess = !isCellValueBlank(cell, resultMap,colEle,row);
		}
		
		return isSuccess;
	}
	/**
	 * 判断单元格中的值是否为空
	 * @param cell
	 * @param resultMap
	 * @return 是否为空的结果
	 */
	public static boolean isCellValueBlank(Cell cell,Map<String, String> resultMap, Element colEle, Row row){
		boolean result = true;
		String cellValue = getCellValue(cell);
		result = StringUtils.isBlank(cellValue);
		if (result) {
			if (cell != null) {
				setFailMsg(cell,resultMap, ExcelConstants.VAL_FAIL_MSG_NOTNULL);
			}else {
				setFailMsg(resultMap, colEle, row,ExcelConstants.VAL_FAIL_MSG_NOTNULL);
			}
			
		}
		return result;
	}

	/**
	 * @param resultMap
	 * @param colEle
	 * @param row
	 */
	private static void setFailMsg(Map<String, String> resultMap, Element colEle, Row row, String failMsg) {
		String fMsg = "行："+(row.getRowNum()+1)+"，列："+colNumToColName(Integer.valueOf(colEle.attributeValue(ExcelConstants.COL_ATTR_COLNUM)))+"，" + failMsg;
		logger.debug("---验证失败---"+fMsg);
		setFailMsg(resultMap, fMsg);
	}

	/**
	 * 正则表达式验证
	 * 
	 * @param cell
	 * @param resultMap
	 * @param valRegExp
	 */
	public static boolean valRegExp(Cell cell, Map<String, String> resultMap, String valRegExp, String failMsg, Row dataRow, Element colEle) {
		boolean isSuccess = true;
		String cellValue = getCellValue(cell);
		boolean isMatch = isRegExpMatch(valRegExp, cellValue);
		if (!isMatch) {
			isSuccess = false;
			if (cell != null) {
				setFailMsg(cell,resultMap, failMsg);
			}else {
				setFailMsg(resultMap, colEle, dataRow, failMsg);
			}
		}
		return isSuccess;
	}

	/**
	 * 设置错误信息追加行列信息
	 * @param cell
	 * @param resultMap
	 * @param failMsg
	 */
	public static void setFailMsg(Cell cell,Map<String, String>resultMap, String failMsg){
		failMsg = "行："+(cell.getRowIndex()+1)+"，列："+ExcelImportUtil.colNumToColName(cell.getColumnIndex()+1)+"，"+failMsg;
		setFailMsg(resultMap, failMsg);
	}
	/**
	 * 在原有的failMsg中加入sheet名称 行数 列数信息
	 * 
	 * @param sheetName
	 * @param rowNum
	 * @param colNum
	 * @param failMsg
	 * @return
	 */
	public static String makeFailMsg(String sheetName, int rowIndex, int colIndex, String failMsg) {
		if (StringUtils.isBlank(sheetName)) {
			return "行：" + (rowIndex + 1) + "， 列：" + ExcelImportUtil.colNumToColName(colIndex + 1) + "，" + failMsg;
		}else {
			return "工作表：" + sheetName + "，行：" + (rowIndex + 1) + "， 列：" + ExcelImportUtil.colNumToColName(colIndex + 1) + "，" + failMsg;
		}
		
	}
	
	/**
	 * 在原有的failMsg中加入sheet名称 行数 列数信息
	 * 
	 * @param sheetName
	 * @param rowNum
	 * @param colNum
	 * @param failMsg
	 * @return
	 */
	public static String makeFailMsg(int rowIndex, int colIndex, String failMsg) {
		return makeFailMsg(null, rowIndex, colIndex, failMsg);
	}

	/**
	 * 给定值与的给定的正则表达式是否匹配
	 * 
	 * @param valRegExp
	 *            给定的正则表达式
	 * @param value
	 *            给定的值
	 * @return
	 */
	private static boolean isRegExpMatch(String valRegExp, String value) {
		Pattern pattern = Pattern.compile(valRegExp);
		boolean isMatch = pattern.matcher(value).matches();
		return isMatch;
	}

	/**
	 * 根据类型组成sql
	 * 
	 * @param sql
	 * @param fieldType
	 *            插入表中的类型
	 * @param cellValue
	 *            单元格的值
	 * @param isInsert
	 *            是否是insert操作
	 * 
	 */
	private static void makeFullSqlByType(Row row, StringBuffer sql, Element colEle, String cellValue,
			Map<String, String> resultMap, Boolean isInsert) throws Exception {
		String tResult = transformIntoSqlFormat(row, colEle, cellValue, resultMap);
		if (StringUtils.isNotBlank(tResult)) {
			String fieldName = colEle.attributeValue(ExcelConstants.COL_ATTR_TABFIELD);// 字段名称
			if (!isInsert) {// update
				sql.append(fieldName);
				sql.append("=");
			}
			sql.append(tResult);
			sql.append(",");
		}
	}

	/**
	 * 将单元格中的值转换为sql形式
	 * 
	 * @param colEle
	 * @param cellValue
	 * @param resultMap
	 */
	private static String transformIntoSqlFormat(Row row, Element colEle, String cellValue, Map<String, String> resultMap) {
		String fieldType = colEle.attributeValue(ExcelConstants.COL_ATTR_FIELDTYPE);// 获得要插入的字段的类型
		String result = "";
		String failMsg = "";
		String rowColInfo = "行："+(row.getRowNum()+1)+"，列："+ExcelImportUtil.colNumToColName(Integer.valueOf(colEle.attributeValue(ExcelConstants.COL_ATTR_COLNUM)))+"，";
		if (ExcelConstants.FIELD_TYPE_DATE.equalsIgnoreCase(fieldType)) {
			if (isRegExpMatch(ExcelConstants.REGEXP_DATE, cellValue)) {
				result = "to_date('" + cellValue + "','" + ExcelConstants.DATE_FORMAT + "')";
			} else {
				failMsg = rowColInfo+cellValue + "的日期格式不为规定的插入日期格式";
				setFailMsg(resultMap, failMsg);
			}

		} else if (ExcelConstants.FIELD_TYPE_NUM.equalsIgnoreCase(fieldType)) {
			if (StringUtils.isNotBlank(cellValue)) {
				if (isRegExpMatch(ExcelConstants.REGEXP_ISNUM, cellValue)) {
					result = cellValue;
				} else {
					failMsg = rowColInfo+"\""+cellValue + "\"应该为数字";
					setFailMsg(resultMap, failMsg);
				}
			} else {// 没有非空数字验证
				Element valEle = (Element) colEle.selectSingleNode("./" + ExcelConstants.ELEMENT_VALIDATION);
				if (valEle == null) {
					result = "";
				} else {
					setFailMsg(resultMap, rowColInfo+"值不能为空");
				}
			}

		} else if (ExcelConstants.FIELD_TYPE_STRING.equalsIgnoreCase(fieldType)) {
			result = "'" + cellValue + "'";
		}
		return result;
	}

	/**
	 * 获得excel单元格的值
	 * 
	 * @param cell
	 * @return
	 */
	public static String getCellValue(Cell cell) {
		String cellValue = "";
		if (cell != null) {
			int cellType = cell.getCellType();// 单元格类型
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
			switch (cellType) {
			case Cell.CELL_TYPE_BLANK:
				cellValue = cell.getRichStringCellValue().getString();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					double dateD = cell.getNumericCellValue();
					Date date = DateUtil.getJavaDate(dateD);
					cellValue = sdf.format(date);
				} else {
					DecimalFormat df = new DecimalFormat("0");
					cellValue = df.format(cell.getNumericCellValue());
				}
				break;
			case Cell.CELL_TYPE_STRING:
				cellValue = cell.getRichStringCellValue().toString().trim();
				break;
			}
		}
		return cellValue;
	}
	
	/**
	 * 检查单元格类型是否为字符，数字或者空，若不为此三种类型均返回false
	 * @param impCell
	 * @param resultMap
	 * @return
	 */
	private static boolean checkCellType(Cell impCell, Map<String, String>resultMap){
		boolean result = true;
		if (impCell != null) {
			int cellType = impCell.getCellType();
			if (Cell.CELL_TYPE_STRING==cellType || Cell.CELL_TYPE_NUMERIC==cellType|| Cell.CELL_TYPE_BLANK==cellType) {
				result = true;
			}else {
				setFailMsg(impCell, resultMap, "该单元格类型错误（支持单元格类型：文本或者数字。其他类型暂不支持）");
				result = false;
			}
		}
		return result;
	}

	/**
	 * 检查文件或目录的路径以及文件是否存在若存在则返回该文件或目录的文件对象
	 * 
	 * @param FilePath
	 *            文件路径
	 * @param isDir
	 *            该路径是否为目录
	 * @throws Exception
	 */
	public static File validateFile(String FilePath, boolean isDir) throws Exception {
		File file = null;
		if (StringUtils.isBlank(FilePath)) {
			logger.error("FilePath is empty");
			throw new Exception("FilePath is empty!");
		}

		file = new File(FilePath);
		if (!file.exists()) {
			logger.error("File：" + FilePath + " is not exist!");
			throw new Exception("File：" + FilePath + " is not exist!");
		}

		if (isDir && !file.isDirectory()) {
			logger.error("File：" + FilePath + " is not Dir!");
			throw new Exception("File：" + FilePath + " is not Dir!");
		}
		return file;
	}

	/**
	 * 生成exce 工作簿对象
	 * 
	 * @param excelFile
	 * @return
	 * @throws Exception
	 */
	public static Workbook genWorkbook(File excelFile, Document document, String templateId,
			Map<String, String> resultMap) throws Exception {
		String fileExt = getFileExt(excelFile.getName());
		long start = System.currentTimeMillis();
		Workbook result = null;
		logger.debug("---excel文件类型---" + fileExt);
		try {
			boolean isProt = isProtected(document, templateId);
			if (ExcelConstants.EXCEL_FILE_TYPE.equalsIgnoreCase(fileExt)) {
				result = genPOIHSSFWorkBook(excelFile, document,templateId, resultMap);
			} else if (ExcelConstants.EXCEL2007_FILE_TYPE.equalsIgnoreCase(fileExt)) {
				if (document != null) {
					result = genPOIXSSFWorkBook(excelFile, isProt, getTemplateElement(document,templateId), resultMap);
				} else {
					result = genPOIXSSFWorkBook(excelFile, resultMap);
				}
			} else {
				setFailMsg(resultMap, "不能解析导入的excel文件");
			}
		} catch (Exception e) {
			logger.warn("不能解析导入的excel文件："+excelFile.getName());
			logger.error(e.getMessage(),e);
			result = null;
		}
		long end = System.currentTimeMillis();
		logger.debug("---获得workbook经历时间---"+(end-start));
		if (result == null) {
			setFailMsg(resultMap, "导入的excel文件不是基于系统提供的模板,请下载本系统提供的最新模板", false);
		}
		return result;
	}

	/**
	 * 是否验证模板的规范性即是否为系统提供模板
	 * @param document
	 * @param templateId
	 * @return
	 */
	private static boolean isProtected(Document document, String templateId){
		boolean result = false;
		Element templateE = getTemplateElement(document, templateId);
		String workBookProtect = templateE.attributeValue(ExcelConstants.TEMPLATE_ATTR_WORKBOOKPROTECT);
		if (workBookProtect != null && !"".equals(workBookProtect) && Boolean.valueOf(workBookProtect)) {
			result = true;
		}
		return result;
	}

	/**
	 * 获得template元素
	 * @param document
	 * @param templateId
	 * @return
	 */
	private static Element getTemplateElement(Document document,
			String templateId) {
		Element root = document.getRootElement();
		Element templateE = (Element) root.selectSingleNode("//template[@id='" + templateId + "']");
		return templateE;
	}
	/**
	 * 获得文件扩展名
	 * 
	 * @param fileName
	 * @return 返回文件扩展名
	 */
	public static String getFileExt(String fileName) {
		String result = "";
		if (fileName != null && !"".equals(fileName)) {
			result = fileName.substring(fileName.lastIndexOf(".") + 1);
		}
		return result;
	}

	/**
	 * 返回Excel2007工作簿对象
	 * 
	 * @param excelFilePath
	 *            文件路径
	 * @param document
	 *            配置文件文档对象
	 * @param templateId
	 *            配置文件中的模板元素的模板id属性
	 * @param resultMap
	 *            结果信息
	 * @return excel工作簿对象
	 */
	public static Workbook genPOIXSSFWorkBook(File excelFile, boolean isProt,Element templateE,
			Map<String, String> resultMap) throws Exception {
		
		FileInputStream fis = null;
		Workbook workbook = null;
		try {
			logger.debug("---excel文件路径---" + excelFile.getAbsolutePath());
			fis = new FileInputStream(excelFile);
			if (isProt) {
				// 若模板中设置了工作簿保护
				try {
					POIFSFileSystem fsys = new POIFSFileSystem(fis);
					EncryptionInfo info = new EncryptionInfo(fsys);
					Decryptor d = Decryptor.getInstance(info);
					if (d.verifyPassword(Decryptor.DEFAULT_PASSWORD)) {
						workbook = WorkbookFactory.create(d.getDataStream(fsys));
					}else {
						setFailMsg(resultMap, "导入的excel文件不是基于系统提供的模板,请下载本系统提供的模板");
						logger.debug("导入的excel文件不是基于系统提供的模板");
					}
				} catch (Exception e) {
					if (e instanceof OfficeXmlFileException) {
						if (fis != null) {
							fis.close();
						}
						fis = new FileInputStream(excelFile);
						workbook = WorkbookFactory.create(fis);
					} else {
						throw e;
					}
				}
			} else {
				// 未设置工作簿保护
				// workbook = new XSSFWorkbook(fis);// 创建可写工作薄
				workbook = WorkbookFactory.create(fis);
			}
			//验证Excel文件是否为基于系统提供的模板
			certificateWorkbook(resultMap, templateE, workbook);
			
		} catch (FileNotFoundException e) {
			throw e;
		} catch (IOException e) {
			throw e;
		} finally {
			if (fis != null) {
				try {
					fis.close();
				} catch (IOException e) {
					throw e;
				}
			}
		}
		if(!ExcelConstants.FAIL.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))){
			resultMap.put(ExcelConstants.RESULT_KEY, ExcelConstants.SUCCESS);
		}
		
		logger.debug("---成功获得workbook---");
		return workbook;
	}

	/**
	 * 验证模板是否为系统提供模板
	 * @param resultMap
	 * @param templateNode
	 * @param workbook
	 */
	private static void certificateWorkbook(Map<String, String> resultMap,
			Element templateNode, Workbook workbook) {
		String certificate = templateNode.attributeValue(ExcelConstants.TEMPLATE_ATTR_CERTIFICATE);
		logger.debug("---配置文件中的模板校验字符串为---"+certificate);
		if (workbook != null && certificate != null && !"".equals(certificate)) {
			if (!certificating(workbook, certificate)) {
				setFailMsg(resultMap, "导入的excel文件不是基于系统提供的模板,请下载本系统提供的模板");
				logger.debug("导入的excel文件不是基于系统提供的模板");
			}
		}
	}

	/**
	 * 返回Excel2007工作簿对象
	 * 
	 * @param excelFilePath
	 *            excel文件路径
	 */
	public static Workbook genPOIXSSFWorkBook(File excelFile, Map<String, String> resultMap) throws Exception {
		FileInputStream fis = null;
		Workbook workbook = null;
		try {
			logger.debug("---excel文件路径---" + excelFile.getAbsolutePath());
			fis = new FileInputStream(excelFile);
			POIFSFileSystem fsys = new POIFSFileSystem(fis);
			EncryptionInfo info = new EncryptionInfo(fsys);
			Decryptor d = Decryptor.getInstance(info);
			if (d.verifyPassword(Decryptor.DEFAULT_PASSWORD)) {
				workbook = WorkbookFactory.create(d.getDataStream(fsys));
			}else {
				
			}

			resultMap.put(ExcelConstants.RESULT_KEY, ExcelConstants.SUCCESS);
		} catch (FileNotFoundException e) {
			resultMap.put(ExcelConstants.RESULT_KEY, ExcelConstants.FAIL);
			resultMap.put(ExcelConstants.MSG_KEY, "文件不存在");
			throw e;
		} catch (IOException e) {
			resultMap.put(ExcelConstants.RESULT_KEY, ExcelConstants.FAIL);
			resultMap.put(ExcelConstants.MSG_KEY, "系统IO错误");
			throw e;
		} catch (Exception e) {
			if (e instanceof OfficeXmlFileException) {
				if (fis != null) {
					fis.close();
				}
				fis = new FileInputStream(excelFile);
				workbook = WorkbookFactory.create(fis);
				resultMap.put(ExcelConstants.RESULT_KEY, ExcelConstants.SUCCESS);
			} else {
				resultMap.put(ExcelConstants.RESULT_KEY, ExcelConstants.FAIL);
				resultMap.put(ExcelConstants.MSG_KEY, "未知错误");
				throw e;
			}
		} finally {
			if (fis != null) {
				try {
					fis.close();
				} catch (IOException e) {
					throw e;
				}
			}
		}
		logger.debug("---成功获得workbook---");
		return workbook;
	}

	/**
	 * 返回Excel97-2003的工作簿对象
	 * 
	 * @param excelFilePath
	 */
	public static Workbook genPOIHSSFWorkBook(File excelFile, Document document,String templateId, Map<String, String>resultMap) throws Exception {
		Workbook workbook = null;
		FileInputStream fis = null;
		try {
			logger.debug("---文件为excel97-2003文件--导入excel文件路径---" + excelFile);
			fis = new FileInputStream(excelFile);
			POIFSFileSystem fsys = new POIFSFileSystem(fis);
			workbook = WorkbookFactory.create(fsys);
			logger.debug("---成功获得workbook---");
			certificateWorkbook(resultMap,getTemplateElement(document, templateId),workbook);
		} catch (FileNotFoundException e) {
			logger.error(e.getMessage(),e);
			throw e;
		} catch (IOException e) {
			logger.error(e.getMessage(),e);
			throw e;
		}catch (Exception e) {
			logger.error(e.getMessage(),e);
			throw e;
		} finally {
			if (fis != null) {
				try {
					fis.close();
				} catch (IOException e) {
					logger.error(e.getMessage(),e);
					throw e;
				}
			}
		}
		if(!ExcelConstants.FAIL.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))){
			resultMap.put(ExcelConstants.RESULT_KEY, ExcelConstants.SUCCESS);
		}
		return workbook;
	}

	/**
	 * 认证工作簿是否为基于系统提供的导入模板
	 * 
	 * @param workbook
	 * @return
	 */
	private static boolean certificating(Workbook workbook, String certificateCode) {
		logger.debug("---开始---方法---certificating");
		boolean result = false;
		Sheet certificateSheet = workbook.getSheet("certificate");
		if (certificateSheet != null) {// 是否存在名为certificate的sheet
			Cell certificateCell = certificateSheet.getRow(0).getCell(0);
			if (certificateCell != null) { // 该认证sheet中的第一行第一列的cell是否为空
				String certificateValue = certificateCell.getRichStringCellValue().getString();
				logger.debug("---解析文件中的校验certificate sheet中的校验字符串为---"+certificateValue);
				if (certificateCode.equals(certificateValue)) {// cell中的值与配置中的值是否相等来验证该文件是否基于系统提供的模板修改
					result = true;
					logger.debug("---配置文件中的校验字符串与解析文件中的校验字符串相同---该文件为标准文件---");
				}
			}
		}
		logger.debug("---结束---方法---certificating");
		return result;
	}

	public static void main(String[] args) {
//		File excelFile = new File("/media/zxchaosExt/virtual_xplarge_share/csgj.xlsx");
//		String configFilePath = "/home/zxchaos/projectFiles/projects/rybt/webapp/eap/WEB-INF/import.xml";
//		String templateId = "csgjImp";
//		// String excelDir = "/home/zxchaos/test/";
//		boolean isValidate = false;
//		try {
//			// Map<String, String> map = ExcelImportUtil.importExcel(excelDir, templateId, configFilePath,
//			// ExcelConstants.EXCEL_FILE_TYPE);
//			Map<String, String> map = importExcel(excelFile, templateId, configFilePath);
//			logger.debug(map.get(ExcelConstants.RESULT_KEY));
//			if (ExcelConstants.FAIL.equalsIgnoreCase(map.get(ExcelConstants.RESULT_KEY))) {
//				System.out.println(map.get(ExcelConstants.MSG_KEY));
//			} else {
//				System.out.println(map.get(ExcelConstants.SQLS_KEY));
//			}
//
//		} catch (Exception e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
		System.out.println(colNumToColName(26));
	}
	
	/**
	 * 将中的列号转换为列名称（A，B...）
	 * @param colNum
	 * @return
	 */
	public static String colNumToColName(int column) {
		int col = column;
		int system = 26;
		char[] digArr = new char[100];
		int ind = 0;
		while (col > 0) {
			int mod = col % system;
			if (mod == 0)
				mod = system;
			digArr[ind++] = dig2Char(mod);
			col = (col - 1) / 26;
		}
		StringBuffer bf = new StringBuffer(ind);
		for (int i = ind - 1; i >= 0; i--) {
			bf.append(digArr[i]);
		}
		return bf.toString();
	}
	    
	    
	private static char dig2Char(final int dig) {
		int acs = dig - 1 + 'A';
		return (char) acs;
	}

}
