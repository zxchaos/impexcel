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
 * Excel ���빤���� ������ɶ��������Excel�ļ��Ĺ�������ϵ��������ļ�import.xml�е����ý��е�������֤
 * 
 * @author zhangxin
 * 
 */
public class ExcelImportUtil {
	public static final Logger logger = LoggerFactory.getLogger(ExcelImportUtil.class);

	/**
	 * ����Excel������ڷ��� ��Ŀ¼impExcelDir�µ����е�excel�ļ����������ļ��е�ͬһ��ģ�嵼��
	 * 
	 * @param impExcelDir
	 *            ���Excel�ļ���Ŀ¼
	 * @param templateId
	 *            Ҫ�����excel�ļ���Ӧ�������ļ���template��ǩ�е�id����ֵ
	 * @param configFilePath
	 *            �����ļ�·��
	 * @param excelFileType
	 *            Ŀ¼�µ�excel�ļ����� ��Ϊ xls �� xlsx ����
	 * 
	 * @return �����е�Map��key����"result":��ʾ�����ɹ���ʧ����ֵΪ"fail"��"success" ������ʧ�ܼ�result��Ӧ��ֵΪfailʱ �����key��"msg" ���а���������Ϣ
	 *         �������ɹ���result��Ӧ��ֵΪsuccessʱ �����key��"sqls" ���а��������ɹ����insert �� update ��� insert �� update ����ʽΪ: 
	 *         INSERT INTO TABLENAME (FIELD1,FIELD2...) @%&#VALUES#&%@ (VALUE1,VALUE2);\n 
	 *         ���� UPDATE TABLENAME SET FIELD1=VALUE1,FIELD2=VALUE2... WHERE FIELD3=VALUE3 AND FIELD4=VALUE4 ...;\n
	 */
	public static Map<String, String> importExcel(String impExcelDir, String templateId, String configFilePath,
			String excelFileType) throws Exception {
		File excelDir = validateFile(impExcelDir, true);
		File[] excelFiles = getDirFiles(excelDir, excelFileType);
		Map<String, String> resultMap = new HashMap<String, String>();// ��Ž��������Map

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
	 * ����Excel������ڷ��� ���������ļ���ģ��id�������ļ�·������Excel
	 * 
	 * @param excelFile
	 *            Ҫ�����excel�ļ�
	 * @param templateId
	 *            ģ��Id
	 * @param configFilePath
	 *            �����ļ�λ��
	 * @return �����е�Map��key����"result":��ʾ�����ɹ���ʧ����ֵΪ"fail"��"success" ������ʧ�ܼ�result��Ӧ��ֵΪfailʱ �����key��"msg" ���а���������Ϣ
	 *         �������ɹ���result��Ӧ��ֵΪsuccessʱ �����key��"sqls" ���а��������ɹ����insert��� insert ����ʽΪ: 
	 *         INSERT INTO TABLENAME (FIELD1,FIELD2...) @%&#VALUES#&%@ (VALUE1,VALUE2);\n 
	 *         ���� 
	 *         UPDATE TABLENAME SET FIELD1=VALUE1,FIELD2=VALUE2...;\n
	 */
	public static Map<String, String> importExcel(File excelFile, String templateId, String configFilePath)
			throws Exception {
		logger.debug("+++��ʼ---����---importExcel");
		Map<String, String> resultMap = new HashMap<String, String>();
		Document document = getConfigFileDoc(configFilePath);
		Workbook workbook = genWorkbook(excelFile, document, templateId, resultMap);
		importExcel(workbook, document, templateId);
		logger.debug("+++����---����---importExcel");
		return resultMap;
	}
	
	
	public static void importExcel(Workbook workbook,Document configDoc, String templateId)
			throws Exception {
		importExcel(workbook, configDoc, templateId, null);
	}
	
	/**
	 * ����Excel������ڷ��� ���������ļ���ģ��id�������ļ�·������Excel
	 * 
	 * @param workbook
	 *            Ҫ�����Ĺ���������
	 * @param configDoc ���������ļ����ĵ�����
	 * @param templateId
	 *            ģ��Id
	 * @param resultMap ��Ž������
	 * @return �����е�Map��key����"result":��ʾ�����ɹ���ʧ����ֵΪ"fail"��"success" ������ʧ�ܼ�result��Ӧ��ֵΪfailʱ �����key��"msg" ���а���������Ϣ
	 *         �������ɹ���result��Ӧ��ֵΪsuccessʱ �����key��"sqls" ���а��������ɹ����insert��� insert ����ʽΪ: 
	 *         INSERT INTO TABLENAME (FIELD1,FIELD2...) @%&#VALUES#&%@ (VALUE1,VALUE2);\n 
	 *         ���� 
	 *         UPDATE TABLENAME SET FIELD1=VALUE1,FIELD2=VALUE2...;\n
	 *         ����
	 *         INSERT INTO TABLENAME (FIELD1,FIELD2...) SELECT VALUE1,VALUE2... FROM DUAL UNION SELECT ...
	 */
	public static Map<String, String> importExcel(Workbook workbook,Document configDoc, String templateId, Map<String, String> resultMap)
			throws Exception {
		logger.debug("+++��ʼ---����---importExcel");
		if (resultMap == null) {
			resultMap = new HashMap<String, String>();
		}
		doImport(workbook, configDoc, templateId,resultMap);
		logger.debug("+++����---����---importExcel");
		return resultMap;
	}

	/**
	 * ����excel����ɹ���Ϣ
	 * 
	 * @param resultMap
	 *            �ܽ��map
	 * @param impMap
	 *            ĳ�ε�����ŵĵ�����Ϣmap
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
	 * ����resultMap��׷��ʧ����Ϣ���ļ���Ϣ
	 * 
	 * @param resultMap
	 *            ���map
	 * @param excelFile
	 *            ������ļ���Ϣ
	 * @param impMap
	 *            �ļ�������
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
			resultMap.put(ExcelConstants.FILENAME_KEY, excelFile.getName());// �����ļ���
		}
		resultMap.remove(ExcelConstants.SQLS_KEY);
	}

	/**
	 * ��resultMap��ʧ����Ϣ׷�ӵ�finalResultMap��
	 * @param finalResultMap
	 * @param resultMap
	 */
	public static void appendFailMsg(Map<String, String> finalResultMap, Map<String, String> resultMap){
		appendFailMsg(finalResultMap, null, resultMap);
	}
	/**
	 * ���Ŀ¼�µ�ָ����չ�����ļ�
	 * 
	 * @param excelDir
	 *            Ҫ���ָ�������ļ���Ŀ¼
	 * @param fileExtName
	 *            Ҫ��õ��ļ����͵���չ��
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
	 * ���Ŀ¼�µĵ�excel�ļ�����Excel 2007+��2003
	 * 
	 * @param excelDir
	 *            Ҫ���excel�ļ���Ŀ¼
	 * @return
	 */
	public static File[] getDirFiles(File excelDir) {
		FilenameFilter filenameFilter = new FilenameFilter() {
			public boolean accept(File dir, String name) {
				boolean result = false;
				Pattern pattern2003 = Pattern.compile("[\\s\\S]*.(" + ExcelConstants.EXCEL_FILE_TYPE + ")$");
				Pattern pattern2007 = Pattern.compile("[\\s\\S]*.(" + ExcelConstants.EXCEL2007_FILE_TYPE + ")$");
				result = pattern2003.matcher(name).matches();
				if (!result) {// ���ļ�����2003 �����ж��Ƿ�Ϊ2007
					result = pattern2007.matcher(name).matches();
				}
				return result;
			}
		};

		File[] excelFiles = excelDir.listFiles(filenameFilter);
		return excelFiles;
	}

	/**
	 * ��õ��������ļ����ĵ�����
	 * 
	 * @param configFilePath
	 * @return �����ļ����ĵ�����
	 * @throws Exception
	 * @throws DocumentException
	 */
	public static Document getConfigFileDoc(String configFilePath) throws Exception, DocumentException {
		long start = System.currentTimeMillis();
		File configFile = validateFile(configFilePath, false);
		SAXReader reader = new SAXReader();
		Document document = reader.read(configFile);
		long end = System.currentTimeMillis();
		logger.debug("---�ɹ���������ļ��ĵ�����---����ʱ��---"+(end-start));
		return document;
	}

	/**
	 * ����Excel ������
	 * 
	 * @param workbook
	 *            ����������
	 * @param doc
	 *            �����ļ���xml�ĵ�����
	 * @param templateId
	 *            �����ļ���ģ���Ӧ��idֵ
	 * @param isValidate
	 *            �Ƿ���е�����֤
	 * @return ����������(ȡ����map�е�keyΪ ExcelConstants.RESULT_KEY
	 *         ��Ӧ��valueֵ{"success","fail"}�����жϵ����Ƿ�ɹ�,��result��ӦvalueΪ"fail"����ʧ�ܣ��ٻ�ȡmap�е�keyΪExcelConstants.MSG_KEY
	 *         ��Ӧ��value�ɻ�õ��������Ϣ)
	 */
	public static void doImport(Workbook workbook, Document doc, String templateId, Map<String, String> resultMap) throws Exception {
		logger.debug("---��ʼ---����---doImport");
		List<Element> sheetEleList = getSheetList(doc, templateId);

		for (Element sheetEle : sheetEleList) {
			String sheetName = sheetEle.attributeValue(ExcelConstants.SHEET_ATTR_NAME);// ��������ļ���template��ǩ�µ�sheet�����Ƽ��������ļ��пɵ����sheet����
			if (sheetName != null && !"".equals(sheetName)) {
				Sheet impSheet = workbook.getSheet(sheetName);// ͨ�����ƻ��workbook�еĽ�Ҫ�����sheet
				if (impSheet != null) {
					impSheet(impSheet, sheetEle,resultMap);
					if (ExcelConstants.FAIL.equalsIgnoreCase(resultMap.get(ExcelConstants.RESULT_KEY))) {
						break;
					}
				} else {
					String failMsg = "���ƣ�" + sheetName + "û�ж�Ӧ��sheet";
					setFailMsg(resultMap, failMsg);
					break;
				}
			}
		}
		logger.debug("+++����---����---doImport");
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
	 * @param isAppend �Ƿ�׷�Ӵ�����Ϣ
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
	 * ���������ļ����ĵ�����������ļ���template��ǩidֵ���ĳһtemplate��ǩ�µ�sheet��ǩ
	 * 
	 * @param doc
	 * @param templateId
	 * @return
	 */
	public static List<Element> getSheetList(Document doc, String templateId) {
		Element root = doc.getRootElement();
		Node templatNode = root.selectSingleNode("//template[@" + ExcelConstants.TEMPLATE_ATTR_ID + "='"
				+ templateId + "']");
		List<Element> sheetEleList = templatNode.selectNodes("./" + ExcelConstants.ELEMENT_SHEET);// Ҫ�����sheet
		return sheetEleList;
	}

	/**
	 * ���������ļ�sheet��ǩ�е����õ���excel�ļ�sheet�е�����
	 * 
	 * @param impSheet
	 * @param sheetEle
	 * @param isValidate
	 *            ����ʱ�Ƿ����У��
	 * @return insert��伯��
	 */
	public static Map<String, String> impSheet(Sheet impSheet, Element sheetEle,Map<String, String>resultMap) throws Exception {
		logger.debug("---��ʼ---����---impSheet");
		StringBuffer genSqls = new StringBuffer();// �����װ��ɵ�insert��update���

		Integer dataStartRowNum = Integer.valueOf(sheetEle
				.attributeValue(ExcelConstants.SHEET_ATTR_DATASTARTROWNUM)) - 1;// ����������ʼ����

		Boolean isValidate = Boolean.valueOf(sheetEle.attributeValue(ExcelConstants.SHEET_ATTR_ISVALIDATE));// �����sheet�Ƿ������֤
		Boolean isWholeRow = Boolean.valueOf(sheetEle.attributeValue(ExcelConstants.SHEET_ATTR_WHOLEROW));//��������Ƿ�Ϊһ����
		
		StringBuffer sql = null;
		String operation = sheetEle.attributeValue(ExcelConstants.SHEET_ATTR_OPERATION);
		boolean isInsert = true;
		int colCount = 0;
		if (operation.equalsIgnoreCase(ExcelConstants.OPERATION_TYPE_INSERT)) {// ����insert����
			sql = makePartInsertSql(sheetEle);
		} else if (operation.equalsIgnoreCase(ExcelConstants.OPERATION_TYPE_UPDATE)) {// ����update����
			sql = makePartUpdateSql(sheetEle);
			isInsert = false;
		} else if (ExcelConstants.OPERATION_TYPE_MULTINSERT.equalsIgnoreCase(operation)) {//multinser����
			sql = makePartMultInsertSql(sheetEle);
			colCount = Integer.valueOf(sheetEle.attributeValue(ExcelConstants.SHEET_ATTR_COLCOUNT));
			isInsert = false;
		} else {
			throw new Exception("sheet ��ǩ��operation�������ô���");
		}

		List<Element> colList = sheetEle.selectNodes("./" + ExcelConstants.ELEMENT_COL);// ���sheetԪ�������е�col��Ԫ��
		List<String> uniqueList = new ArrayList<String>();
		
		for (int i = dataStartRowNum; i <= impSheet.getLastRowNum(); i++) {
			StringBuffer rowSql = new StringBuffer(sql);// �����sheet��ÿһ�����һ��insert ���� update ���
			Map<String, String> rowResultMap = new HashMap<String, String>();
			Row dataRow = impSheet.getRow(i);
			if(dataRow == null){
				logger.warn("+++�У�"+(i+1)+"Ϊ��+++");
				continue;
			}
			logger.debug("---������---" + (dataRow.getRowNum()+1));
			
			StringBuffer updateWhere = new StringBuffer("");// update����������
			int nullColNum = 0;//ֵΪ�յ�����
			int multInsertCount = 0;//multinsert����ʱ����
			int multInsertBlankNum = 0;//multinsert����ʱ�Ŀ��е���Ŀ
			for (Element colEle : colList) {
				Integer colNum = Integer.valueOf(colEle.attributeValue(ExcelConstants.COL_ATTR_COLNUM).trim()) - 1;// ���Ҫ������е��к�
				
				Cell impCell = dataRow.getCell(colNum);
				
				if (!checkCellType(impCell, rowResultMap)) {
					continue;
				}
				
				//Ҫ��������Ƿ�Ϊ��
				if (null == impCell || StringUtils.isBlank(getCellValue(impCell))) {
					nullColNum++;
				}
				
				if (isValidate != null && isValidate 
						&& !validating(rowResultMap, colEle, impCell, impSheet, dataRow)) {// �����sheet��Ҫ��֤
						continue;
				}
				//���������ʽ
				if (!regExpCheck(colEle, impCell,rowResultMap,dataRow)) {
					continue;
				}
				
				//���excel�е��ظ�ֵ
				if (checkUnique(rowResultMap, uniqueList, colEle, impCell)) {
					continue;
				}
				
				//�������ֵ����󳤶�
				if (!checkLength(rowResultMap, colEle, impCell)) {
					continue;
				}
				Boolean isCondition = Boolean
						.valueOf(colEle.attributeValue(ExcelConstants.COL_ATTR_ISCONDITION));
				if (isCondition != null && isCondition) {// �õ�Ԫ����Ϊupdate����
					makeUpdateWhere(dataRow,updateWhere, colEle, impCell, rowResultMap);
				} else if (ExcelConstants.OPERATION_TYPE_MULTINSERT.equalsIgnoreCase(operation)) {//������Ϊmultinsert
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
							logger.debug("---multinsert---�����ֵ��Ϊ��---��ʱ��insertsqlΪ---"+rowSql);
							rowSql.replace(rowSql.lastIndexOf(ExcelConstants.SQL_MULTINSERT_UNION_SELECT_FLAG)+ExcelConstants.SQL_MULTINSERT_UNION_SELECT_FLAG.length(), rowSql.length(), "");
							logger.debug("---multinsert---ȥ��ȫ��Ϊ��һ����sql---"+rowSql);
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
			logger.debug("��"+(i+1)+"�Ŀ�����ĿΪ"+nullColNum+"---���õ��е���ĿΪ"+configColNum+"---�����е�wholeRow����ֵΪ��"+isWholeRow);
			
			if (null != isWholeRow && isWholeRow && nullColNum == configColNum) {//��һ���ж�Ϊ��
				logger.info("�У�"+(i+1)+"Ϊ��");
				continue;
			}else if (ExcelConstants.FAIL.equals(rowResultMap.get(ExcelConstants.RESULT_KEY))) {
				// У��ʧ��
					appendFailMsg(resultMap, rowResultMap);
					continue;
			}
			
			if (ExcelConstants.FAIL.equals(resultMap.get(ExcelConstants.RESULT_KEY))) {// У��ʧ��
				continue;
			}
			if (ExcelConstants.OPERATION_TYPE_INSERT.equals(operation)) {// insert����
				rowSql.replace(rowSql.lastIndexOf(","), rowSql.length(), ")");
			} else if(ExcelConstants.OPERATION_TYPE_UPDATE.equals(operation)) {// update ����
				makeUpdateRowSql(rowSql, updateWhere);
				if (!StringUtils.contains(rowSql, "WHERE")) {// �Ƿ����where
					String failMsg = "���ɵ�update��䲻����where�Ӿ䣡";
					logger.error(failMsg);
					setFailMsg(resultMap, failMsg);
					break;
				}
			}else if (ExcelConstants.OPERATION_TYPE_MULTINSERT.equals(operation)) {//һ��insert���������¼
				rowSql.replace(rowSql.lastIndexOf(ExcelConstants.SQL_MULTINSERT_UNION_SELECT_FLAG), rowSql.length(),"");
				logger.debug("---multinsert---���д����������sql---"+rowSql);
			}
			
			genSqls.append(rowSql);
			genSqls.append(ExcelConstants.SQL_TAIL);
		}

		if (!ExcelConstants.FAIL.equals(resultMap.get(ExcelConstants.RESULT_KEY))) {
			resultMap.put(ExcelConstants.RESULT_KEY, ExcelConstants.SUCCESS);
			resultMap.put(ExcelConstants.SQLS_KEY, genSqls.toString());
		}
		logger.debug("+++����---����---impSheet");
		return resultMap;
	}

	/**
	 * �����ַ��������Ƿ񳬳���󳤶�
	 * @param resultMap
	 * @param colEle
	 * @param impCell
	 * @return true:��֤�ɹ���false����֤ʧ��
	 */
	private static boolean checkLength(Map<String, String> resultMap, Element colEle, Cell impCell) {
		boolean result = true;
		String maxLength = colEle.attributeValue(ExcelConstants.COL_ATTR_MAXLENGTH);
		if (StringUtils.isNotBlank(maxLength)) {
			int mLength = Integer.valueOf(maxLength);
			String cellValue = getCellValue(impCell);
			//����ƥ��ģʽ
			Pattern zhCharPattern = Pattern.compile("[\u4E00-\u9FA5\uFF00-\uFFEF\u3000-\u303F]");
			Matcher zhMatcher = zhCharPattern.matcher(cellValue);
			int length = cellValue.length();
			while (zhMatcher.find()) {
				//��Ϊ�������ݿ�ʱһ�����ֵ��������ַ�����Ҫ��1
				length++;
			}
			if (length > mLength) {
				setFailMsg(impCell, resultMap, "ֵ�ĳ��ȳ�������");
				result = false;
			}
		}
		return result;
	}
	/**
	 * ���ĳ���е�ֵ��excel�������Ƿ����ظ�
	 * @param resultMap
	 * @param uniqueList
	 * @param colEle
	 * @param impCell
	 * @return true:�����ظ���false���������ظ�
	 */
	private static boolean checkUnique(Map<String, String> resultMap, List<String> uniqueList, Element colEle,
			Cell impCell) {
		boolean result = false;
		String unique = colEle.attributeValue(ExcelConstants.COL_ATTR_UNIQUE);
		if (StringUtils.isNotBlank(unique) && Boolean.valueOf(unique)) {
			String cellValue = getCellValue(impCell);
			if (uniqueList.contains(cellValue)) {
				setFailMsg(impCell, resultMap, "'"+cellValue+"'�ڸ������Ѵ����ظ�ֵ");
				result = true;
			}else {
				uniqueList.add(cellValue);
				result = false;
			}
		}
		return result;
	}
	

	/**
	 * ��鵼��ĵ�Ԫ���е�ֵ�Ƿ��������������ʽ�����ƥ��
	 * @param colEle
	 * @param impCell
	 * @return ƥ����
	 * @throws ClassNotFoundException 
	 * @throws IllegalAccessException 
	 * @throws InstantiationException 
	 */
	private static boolean regExpCheck(Element colEle, Cell impCell, Map<String, String>resultMap, Row dataRow) throws ClassNotFoundException, InstantiationException, IllegalAccessException {
		Element validation = (Element) colEle.selectSingleNode("./" + ExcelConstants.ELEMENT_VALIDATION);
		boolean hasValidation = hasVal(validation);
		
		boolean result = true;
		String recheckerAttrVal = colEle.attributeValue(ExcelConstants.COL_ATTR_REGEXPCHECKER);
		if (StringUtils.isNotBlank(recheckerAttrVal)) {//�Ƿ�����������ʽ����� ���� ����֤����
			IRegExpChecker regExpChecker = (IRegExpChecker)Class.forName(recheckerAttrVal).newInstance();
			if (impCell != null) {
				String cellValue = getCellValue(impCell);
				if (StringUtils.isNotBlank(cellValue)) {
					result = regExpChecker.check(cellValue);
					if (!result) {
						setFailMsg(impCell, resultMap, "����дֵ�ĸ�ʽ����");
						logger.debug("---������ʽ����ʧ��---");
					}
				}else {
					if (hasValidation) {
						setFailMsg(impCell,resultMap, "�õ�Ԫ����Ϊ��");
						result = false;
					}
				}
			}else {
				if (hasValidation) {
					setFailMsg(resultMap, colEle, dataRow, "�õ�Ԫ����Ϊ��");
					result = false;
				}
			}
		}
		return result;
	}
	
	/**
	 * ����Ҫ���µ��е� update sql���
	 * 
	 * @param rowSql
	 * @param updateWhere
	 */
	private static void makeUpdateRowSql(StringBuffer rowSql, StringBuffer updateWhere) {
		if (StringUtils.isNotBlank(updateWhere)) {// ��where�Ӿ䲻Ϊ��
			// ��where�Ӿ����Ķ����ANDȥ��
			updateWhere.replace(updateWhere.lastIndexOf(" AND "), updateWhere.length(), "");
			String where = " WHERE " + updateWhere;
			rowSql.replace(rowSql.lastIndexOf(","), rowSql.length(), where);
		} else {
			rowSql.replace(rowSql.lastIndexOf(","), rowSql.length(), "");
		}
	}
	
	/**
	 * ����multinsert���
	 * @param row
	 * @param multInsertSql
	 * @param colEle
	 * @param cell
	 * @param resultMap
	 * @param isRecord �Ƿ�Ϊ������һ����¼
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
	 * ����update������where�����Ӿ� ���ɵ�where�Ӿ��ʽΪ field1=value1 AND field1=value1 AND ...
	 * 
	 * @param updateWhere
	 * @param colEle
	 *            ���е�����
	 * @param whereCell
	 *            ��Ϊ�����ĵ�Ԫ�����
	 */
	private static void makeUpdateWhere(Row row,StringBuffer updateWhere, Element colEle, Cell whereCell,
			Map<String, String> resultMap) {
		String cellValue = getCellValue(whereCell);
		if (StringUtils.isNotBlank(cellValue.trim())) {// ��cellֵ��Ϊ��������where����
			String fieldName = colEle.attributeValue(ExcelConstants.COL_ATTR_TABFIELD);
			updateWhere.append(fieldName);
			updateWhere.append("=");
			updateWhere.append(transformIntoSqlFormat(row, colEle, cellValue, resultMap));
			updateWhere.append(" AND ");
		} else {
			String failMsg = "�У�"+(row.getRowNum()+1)+"����Ϊ���������ĵ�Ԫ���ֵ����Ϊ��";
			logger.warn(failMsg);
			setFailMsg(resultMap, failMsg);
		}
	}

	/**
	 * ���ÿһ��Ҫ��insert �� update sql
	 * 
	 * @param rowSql
	 *            ��ƴ��ɵĲ���insert �� update sql
	 * @param colEle
	 *            �����ļ��е�colԪ��
	 * @param impCell
	 *            �����excel�ļ��еĵ�Ԫ��
	 */
	private static void makeRowSql(Row row, StringBuffer rowSql, Element colEle, Cell impCell, Map<String, String> resultMap,
			Boolean isInsert) throws Exception {
		
		String cellValue = getCellValueByConvertion(colEle, impCell,resultMap,row);
		makeFullSqlByType(row, rowSql, colEle, cellValue, resultMap, isInsert);
	}
	
	/**
	 * ����ת������õ�Ԫ���ֵ
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

		if (converterVal != null && !"".equals(converterVal)) {// ����������ת����
			if (StringUtils.isBlank(cellValue)) {
				String failMsg = "���в���Ϊ��";
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
				setFailMsg(impCell, resultMap, "�õ�Ԫ��ֵ����д�����Ϲ淶");
			}
		}
		return cellValue;
	}

	/**
	 * ��֤����
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
	 * �Ƿ�������֤Ԫ��
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
	 * ���������ļ��е�sheetԪ�������insert���ģ�values ֮ǰ�Ĳ���
	 * 
	 * @param sheetEle
	 * @return
	 */
	private static StringBuffer makePartInsertSql(Element sheetEle) {
		List<Element> colList = sheetEle.selectNodes("./" + ExcelConstants.ELEMENT_COL);// ���sheetԪ�������е�col��Ԫ��
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
	 * ���ɲ��ֵ�update���
	 * 
	 * @param sheetElement
	 *            �����ļ��е�sheet��ǩ
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
	 * ���ɲ��ֶ��в����insert sql ����INSERT INTO TABLENAME (FIELD1,FIELD2...) SELECT
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
		logger.debug("---���ֶ��в���insert sql����---"+result);
		return result;
	}

	/**
	 * ��֤excel��Ԫ��
	 * 
	 * @param cell
	 *            Ҫ��֤��excel��Ԫ��
	 * @param validateRule
	 *            ��֤����
	 * @param colEle ������Ԫ��
	 * @param row �������
	 * @return ��֤�Ƿ�ͨ��
	 */
	public static boolean doValidate(Cell cell, String valType,Map<String, String>resultMap, Element colEle, Row row) {
		boolean isSuccess = true;
		if (valType.equalsIgnoreCase(ExcelConstants.VALIDATION_ISNUM)) {// ��֤�Ƿ�����
			isSuccess = valRegExp(cell, resultMap, ExcelConstants.REGEXP_ISNUM, ExcelConstants.VAL_FAIL_MSG_ISNUM, row, colEle);
		} else if (valType.equalsIgnoreCase(ExcelConstants.VALIDATION_IS_POSITIVE_NUM)) {//������������֤
			isSuccess = valRegExp(cell, resultMap, ExcelConstants.REGEXP_IS_POSITIVE_NUM, ExcelConstants.VAL_FAIL_MSG_ISPOSITIVENUM, row, colEle);
		}else if (valType.equalsIgnoreCase(ExcelConstants.VALIDATION_NOTNULL)) {// ��֤�Ƿ�Ϊ��
			isSuccess = !isCellValueBlank(cell, resultMap,colEle,row);
		}
		
		return isSuccess;
	}
	/**
	 * �жϵ�Ԫ���е�ֵ�Ƿ�Ϊ��
	 * @param cell
	 * @param resultMap
	 * @return �Ƿ�Ϊ�յĽ��
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
		String fMsg = "�У�"+(row.getRowNum()+1)+"���У�"+colNumToColName(Integer.valueOf(colEle.attributeValue(ExcelConstants.COL_ATTR_COLNUM)))+"��" + failMsg;
		logger.debug("---��֤ʧ��---"+fMsg);
		setFailMsg(resultMap, fMsg);
	}

	/**
	 * ������ʽ��֤
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
	 * ���ô�����Ϣ׷��������Ϣ
	 * @param cell
	 * @param resultMap
	 * @param failMsg
	 */
	public static void setFailMsg(Cell cell,Map<String, String>resultMap, String failMsg){
		failMsg = "�У�"+(cell.getRowIndex()+1)+"���У�"+ExcelImportUtil.colNumToColName(cell.getColumnIndex()+1)+"��"+failMsg;
		setFailMsg(resultMap, failMsg);
	}
	/**
	 * ��ԭ�е�failMsg�м���sheet���� ���� ������Ϣ
	 * 
	 * @param sheetName
	 * @param rowNum
	 * @param colNum
	 * @param failMsg
	 * @return
	 */
	public static String makeFailMsg(String sheetName, int rowIndex, int colIndex, String failMsg) {
		if (StringUtils.isBlank(sheetName)) {
			return "�У�" + (rowIndex + 1) + "�� �У�" + ExcelImportUtil.colNumToColName(colIndex + 1) + "��" + failMsg;
		}else {
			return "������" + sheetName + "���У�" + (rowIndex + 1) + "�� �У�" + ExcelImportUtil.colNumToColName(colIndex + 1) + "��" + failMsg;
		}
		
	}
	
	/**
	 * ��ԭ�е�failMsg�м���sheet���� ���� ������Ϣ
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
	 * ����ֵ��ĸ�����������ʽ�Ƿ�ƥ��
	 * 
	 * @param valRegExp
	 *            ������������ʽ
	 * @param value
	 *            ������ֵ
	 * @return
	 */
	private static boolean isRegExpMatch(String valRegExp, String value) {
		Pattern pattern = Pattern.compile(valRegExp);
		boolean isMatch = pattern.matcher(value).matches();
		return isMatch;
	}

	/**
	 * �����������sql
	 * 
	 * @param sql
	 * @param fieldType
	 *            ������е�����
	 * @param cellValue
	 *            ��Ԫ���ֵ
	 * @param isInsert
	 *            �Ƿ���insert����
	 * 
	 */
	private static void makeFullSqlByType(Row row, StringBuffer sql, Element colEle, String cellValue,
			Map<String, String> resultMap, Boolean isInsert) throws Exception {
		String tResult = transformIntoSqlFormat(row, colEle, cellValue, resultMap);
		if (StringUtils.isNotBlank(tResult)) {
			String fieldName = colEle.attributeValue(ExcelConstants.COL_ATTR_TABFIELD);// �ֶ�����
			if (!isInsert) {// update
				sql.append(fieldName);
				sql.append("=");
			}
			sql.append(tResult);
			sql.append(",");
		}
	}

	/**
	 * ����Ԫ���е�ֵת��Ϊsql��ʽ
	 * 
	 * @param colEle
	 * @param cellValue
	 * @param resultMap
	 */
	private static String transformIntoSqlFormat(Row row, Element colEle, String cellValue, Map<String, String> resultMap) {
		String fieldType = colEle.attributeValue(ExcelConstants.COL_ATTR_FIELDTYPE);// ���Ҫ������ֶε�����
		String result = "";
		String failMsg = "";
		String rowColInfo = "�У�"+(row.getRowNum()+1)+"���У�"+ExcelImportUtil.colNumToColName(Integer.valueOf(colEle.attributeValue(ExcelConstants.COL_ATTR_COLNUM)))+"��";
		if (ExcelConstants.FIELD_TYPE_DATE.equalsIgnoreCase(fieldType)) {
			if (isRegExpMatch(ExcelConstants.REGEXP_DATE, cellValue)) {
				result = "to_date('" + cellValue + "','" + ExcelConstants.DATE_FORMAT + "')";
			} else {
				failMsg = rowColInfo+cellValue + "�����ڸ�ʽ��Ϊ�涨�Ĳ������ڸ�ʽ";
				setFailMsg(resultMap, failMsg);
			}

		} else if (ExcelConstants.FIELD_TYPE_NUM.equalsIgnoreCase(fieldType)) {
			if (StringUtils.isNotBlank(cellValue)) {
				if (isRegExpMatch(ExcelConstants.REGEXP_ISNUM, cellValue)) {
					result = cellValue;
				} else {
					failMsg = rowColInfo+"\""+cellValue + "\"Ӧ��Ϊ����";
					setFailMsg(resultMap, failMsg);
				}
			} else {// û�зǿ�������֤
				Element valEle = (Element) colEle.selectSingleNode("./" + ExcelConstants.ELEMENT_VALIDATION);
				if (valEle == null) {
					result = "";
				} else {
					setFailMsg(resultMap, rowColInfo+"ֵ����Ϊ��");
				}
			}

		} else if (ExcelConstants.FIELD_TYPE_STRING.equalsIgnoreCase(fieldType)) {
			result = "'" + cellValue + "'";
		}
		return result;
	}

	/**
	 * ���excel��Ԫ���ֵ
	 * 
	 * @param cell
	 * @return
	 */
	public static String getCellValue(Cell cell) {
		String cellValue = "";
		if (cell != null) {
			int cellType = cell.getCellType();// ��Ԫ������
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
	 * ��鵥Ԫ�������Ƿ�Ϊ�ַ������ֻ��߿գ�����Ϊ���������;�����false
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
				setFailMsg(impCell, resultMap, "�õ�Ԫ�����ʹ���֧�ֵ�Ԫ�����ͣ��ı��������֡����������ݲ�֧�֣�");
				result = false;
			}
		}
		return result;
	}

	/**
	 * ����ļ���Ŀ¼��·���Լ��ļ��Ƿ�����������򷵻ظ��ļ���Ŀ¼���ļ�����
	 * 
	 * @param FilePath
	 *            �ļ�·��
	 * @param isDir
	 *            ��·���Ƿ�ΪĿ¼
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
			logger.error("File��" + FilePath + " is not exist!");
			throw new Exception("File��" + FilePath + " is not exist!");
		}

		if (isDir && !file.isDirectory()) {
			logger.error("File��" + FilePath + " is not Dir!");
			throw new Exception("File��" + FilePath + " is not Dir!");
		}
		return file;
	}

	/**
	 * ����exce ����������
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
		logger.debug("---excel�ļ�����---" + fileExt);
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
				setFailMsg(resultMap, "���ܽ��������excel�ļ�");
			}
		} catch (Exception e) {
			logger.warn("���ܽ��������excel�ļ���"+excelFile.getName());
			logger.error(e.getMessage(),e);
			result = null;
		}
		long end = System.currentTimeMillis();
		logger.debug("---���workbook����ʱ��---"+(end-start));
		if (result == null) {
			setFailMsg(resultMap, "�����excel�ļ����ǻ���ϵͳ�ṩ��ģ��,�����ر�ϵͳ�ṩ������ģ��", false);
		}
		return result;
	}

	/**
	 * �Ƿ���֤ģ��Ĺ淶�Լ��Ƿ�Ϊϵͳ�ṩģ��
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
	 * ���templateԪ��
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
	 * ����ļ���չ��
	 * 
	 * @param fileName
	 * @return �����ļ���չ��
	 */
	public static String getFileExt(String fileName) {
		String result = "";
		if (fileName != null && !"".equals(fileName)) {
			result = fileName.substring(fileName.lastIndexOf(".") + 1);
		}
		return result;
	}

	/**
	 * ����Excel2007����������
	 * 
	 * @param excelFilePath
	 *            �ļ�·��
	 * @param document
	 *            �����ļ��ĵ�����
	 * @param templateId
	 *            �����ļ��е�ģ��Ԫ�ص�ģ��id����
	 * @param resultMap
	 *            �����Ϣ
	 * @return excel����������
	 */
	public static Workbook genPOIXSSFWorkBook(File excelFile, boolean isProt,Element templateE,
			Map<String, String> resultMap) throws Exception {
		
		FileInputStream fis = null;
		Workbook workbook = null;
		try {
			logger.debug("---excel�ļ�·��---" + excelFile.getAbsolutePath());
			fis = new FileInputStream(excelFile);
			if (isProt) {
				// ��ģ���������˹���������
				try {
					POIFSFileSystem fsys = new POIFSFileSystem(fis);
					EncryptionInfo info = new EncryptionInfo(fsys);
					Decryptor d = Decryptor.getInstance(info);
					if (d.verifyPassword(Decryptor.DEFAULT_PASSWORD)) {
						workbook = WorkbookFactory.create(d.getDataStream(fsys));
					}else {
						setFailMsg(resultMap, "�����excel�ļ����ǻ���ϵͳ�ṩ��ģ��,�����ر�ϵͳ�ṩ��ģ��");
						logger.debug("�����excel�ļ����ǻ���ϵͳ�ṩ��ģ��");
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
				// δ���ù���������
				// workbook = new XSSFWorkbook(fis);// ������д������
				workbook = WorkbookFactory.create(fis);
			}
			//��֤Excel�ļ��Ƿ�Ϊ����ϵͳ�ṩ��ģ��
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
		
		logger.debug("---�ɹ����workbook---");
		return workbook;
	}

	/**
	 * ��֤ģ���Ƿ�Ϊϵͳ�ṩģ��
	 * @param resultMap
	 * @param templateNode
	 * @param workbook
	 */
	private static void certificateWorkbook(Map<String, String> resultMap,
			Element templateNode, Workbook workbook) {
		String certificate = templateNode.attributeValue(ExcelConstants.TEMPLATE_ATTR_CERTIFICATE);
		logger.debug("---�����ļ��е�ģ��У���ַ���Ϊ---"+certificate);
		if (workbook != null && certificate != null && !"".equals(certificate)) {
			if (!certificating(workbook, certificate)) {
				setFailMsg(resultMap, "�����excel�ļ����ǻ���ϵͳ�ṩ��ģ��,�����ر�ϵͳ�ṩ��ģ��");
				logger.debug("�����excel�ļ����ǻ���ϵͳ�ṩ��ģ��");
			}
		}
	}

	/**
	 * ����Excel2007����������
	 * 
	 * @param excelFilePath
	 *            excel�ļ�·��
	 */
	public static Workbook genPOIXSSFWorkBook(File excelFile, Map<String, String> resultMap) throws Exception {
		FileInputStream fis = null;
		Workbook workbook = null;
		try {
			logger.debug("---excel�ļ�·��---" + excelFile.getAbsolutePath());
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
			resultMap.put(ExcelConstants.MSG_KEY, "�ļ�������");
			throw e;
		} catch (IOException e) {
			resultMap.put(ExcelConstants.RESULT_KEY, ExcelConstants.FAIL);
			resultMap.put(ExcelConstants.MSG_KEY, "ϵͳIO����");
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
				resultMap.put(ExcelConstants.MSG_KEY, "δ֪����");
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
		logger.debug("---�ɹ����workbook---");
		return workbook;
	}

	/**
	 * ����Excel97-2003�Ĺ���������
	 * 
	 * @param excelFilePath
	 */
	public static Workbook genPOIHSSFWorkBook(File excelFile, Document document,String templateId, Map<String, String>resultMap) throws Exception {
		Workbook workbook = null;
		FileInputStream fis = null;
		try {
			logger.debug("---�ļ�Ϊexcel97-2003�ļ�--����excel�ļ�·��---" + excelFile);
			fis = new FileInputStream(excelFile);
			POIFSFileSystem fsys = new POIFSFileSystem(fis);
			workbook = WorkbookFactory.create(fsys);
			logger.debug("---�ɹ����workbook---");
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
	 * ��֤�������Ƿ�Ϊ����ϵͳ�ṩ�ĵ���ģ��
	 * 
	 * @param workbook
	 * @return
	 */
	private static boolean certificating(Workbook workbook, String certificateCode) {
		logger.debug("---��ʼ---����---certificating");
		boolean result = false;
		Sheet certificateSheet = workbook.getSheet("certificate");
		if (certificateSheet != null) {// �Ƿ������Ϊcertificate��sheet
			Cell certificateCell = certificateSheet.getRow(0).getCell(0);
			if (certificateCell != null) { // ����֤sheet�еĵ�һ�е�һ�е�cell�Ƿ�Ϊ��
				String certificateValue = certificateCell.getRichStringCellValue().getString();
				logger.debug("---�����ļ��е�У��certificate sheet�е�У���ַ���Ϊ---"+certificateValue);
				if (certificateCode.equals(certificateValue)) {// cell�е�ֵ�������е�ֵ�Ƿ��������֤���ļ��Ƿ����ϵͳ�ṩ��ģ���޸�
					result = true;
					logger.debug("---�����ļ��е�У���ַ���������ļ��е�У���ַ�����ͬ---���ļ�Ϊ��׼�ļ�---");
				}
			}
		}
		logger.debug("---����---����---certificating");
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
	 * ���е��к�ת��Ϊ�����ƣ�A��B...��
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
