package top.tangshitai.excel.utils;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.text.DateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import top.tangshitai.excel.App;
import top.tangshitai.excel.bean.POIResult;
import top.tangshitai.excel.exception.POIException;

public class POIUtils {
	public static List<POIResult> analysisWorkBook(String prefix,String fileName,String[]... sheetConfig) throws POIException {
		try {
			InputStream in = new FileInputStream(path(prefix,fileName));
			ByteArrayInputStream byteIn = new ByteArrayInputStream(readToByte(in));
			Workbook workbook = WorkbookFactory.create(byteIn);
			List<POIResult> result = analysisSheet(workbook, sheetConfig);
			return result;
		}catch (POIException e) {
			throw e;
		}catch (Exception e) {
			throw new POIException("解析excel["+fileName+"]出错",e);
		}
		
	}
	public static List<POIResult> analysisWorkBook(InputStream in,String[]... sheetConfig) throws POIException {
		try {
			ByteArrayInputStream byteIn = new ByteArrayInputStream(readToByte(in));
			Workbook workbook = WorkbookFactory.create(byteIn);
			List<POIResult> result = analysisSheet(workbook, sheetConfig);
			return result;
		}catch (POIException e) {
			throw e;
		}catch (Exception e) {
			throw new POIException("解析excel出错",e);
		}
		
	}
	private static List<POIResult> analysisSheet(Workbook workbook ,String[][] sheetCfg) throws POIException {
		List<POIResult> poirs = new ArrayList<>();
		if(sheetCfg==null||sheetCfg.length<1) new POIException("excel的每个sheet的head没有配置");
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		for(int sheetIndex=0;sheetIndex<workbook.getNumberOfSheets();sheetIndex++) {
			Sheet sheet = workbook.getSheetAt(sheetIndex);
			if(sheetCfg.length==1) {
				poirs.add(analysisRow(sheetCfg[0], sheet, evaluator));
			}else {
				poirs.add(analysisRow(sheetCfg[sheetIndex], sheet, evaluator));
			}
		}
		return poirs;
	}
	private static POIResult analysisRow(String[] headParams,Sheet sheet,FormulaEvaluator evaluator) throws POIException {
		int cellLength = headParams.length;//col长度
		POIResult poiResult = new POIResult(sheet.getSheetName(),cellLength);
		for(int rowIndex=0;rowIndex<sheet.getLastRowNum();rowIndex++) {
			Row row = sheet.getRow(rowIndex);
			String[] rowResult = analysisCell(cellLength, row, evaluator);
			if(rowIndex==0) {//处理标题
				boolean flag = false;
				StringBuilder sb = new StringBuilder();
				for(int i1=0;i1<cellLength;i1++) {
					if(!headParams[i1].equals(rowResult[i1])) {
						flag = true;
						sb.append("{["+headParams[i1]+"]和["+rowResult[i1]+"]请修正为-->["+headParams[i1]+"]},");
					}else {
						sb.append("{"+headParams[i1]+"},");
					}
				}
				if(flag)  throw new POIException("excel标题与配置的标题不同，请修改excel标题为:"+sb.toString());
				continue;
			}
			poiResult.add(rowResult);
		}
		return poiResult;
	}
	private static String[] analysisCell(int cellLength,Row row,FormulaEvaluator evaluator) throws POIException {
		String[] cellsContent = new String[cellLength];
		for(int cellIndex=0;cellIndex<cellLength;cellIndex++) {
			Cell cell = row.getCell(cellIndex);
			String cellResult = getCellStrValue(cell,evaluator);
			cellsContent[cellIndex]=StringUtils.isBlank(cellResult)?"":cellResult.trim();
		}
		return cellsContent;
	}
	private static String getCellStrValue(Cell cell,FormulaEvaluator evaluator) throws POIException {
		if(cell==null) throw new IllegalArgumentException("poi的cell对象不能为空");
		switch (cell.getCellType()) {
		case _NONE://不存在类型
			throw new POIException("poi的cell type不存在");
		case NUMERIC://数值型，日期型(excel日期用数值型表示)
	        if (DateUtil.isCellDateFormatted(cell)) {
	        	Date dateTmp = cell.getDateCellValue();
	        	if(dateTmp == null) return ""; 
	        	return DateFormat.getDateTimeInstance().format(cell.getDateCellValue());
	        } else {
	        	String doubleStr = String.valueOf(cell.getNumericCellValue());
	        	if(StringUtils.isBlank(doubleStr)) return "";
	            if (doubleStr.contains("E")) {
	            	int indexOfPoint = doubleStr.indexOf('.');
	                int indexOfE = doubleStr.indexOf('E');
	                // 小数部分
	                BigInteger xs = new BigInteger(doubleStr.substring(indexOfPoint + BigInteger.ONE.intValue(), indexOfE));
	                // 指数
	                int pow = Integer.valueOf(doubleStr.substring(indexOfE + BigInteger.ONE.intValue()));
	                int xsLen = xs.toByteArray().length;
	                int scale = xsLen - pow > 0 ? xsLen - pow : 0;
	                doubleStr = String.format("%." + scale + "f", cell.getNumericCellValue());
	            } else {
	                java.util.regex.Pattern p = Pattern.compile(".0$");
	                java.util.regex.Matcher m = p.matcher(doubleStr);
	                if (m.find()) {
	                    doubleStr = doubleStr.replace(".0", "");
	                }
	            }
	            return doubleStr;
	        }
		case STRING://字符型
			return StringUtils.trimToEmpty(cell.getStringCellValue());
		case FORMULA://计算型
			return String.valueOf(evaluator.evaluate(cell).getNumberValue());
		case BLANK://空白型
			return "";
		case BOOLEAN://boolean型
			return String.valueOf(cell.getBooleanCellValue());
		case ERROR://错误型
			throw new POIException("poi的cell type是错误类型");
		default:
			throw new POIException("poi的不支持的cell type");
		}
	}
	private static String path(String prefix,String filename) throws POIException {
		if(!StringUtils.isBlank(prefix)) {
			if(prefix.indexOf(prefix.length()-1)!='/') {
				filename = prefix+"/"+filename;
			}else {
				filename = prefix+filename;
			}
		}
		java.net.URL url = App.class.getClassLoader().getResource(filename);
		if(url!=null) {
			return url.getPath();
		}
		File absPathFile = new File(filename);
		if(!absPathFile.exists())
			throw new POIException("加载的配置文件不存在："+filename);
		return absPathFile.getAbsolutePath();
	}
	private static byte[] readToByte(InputStream inputStream) throws POIException{
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            byte[] buffer = new byte[1024];
            int num = inputStream.read(buffer);
            while (num != -1) {
                baos.write(buffer, 0, num);
                num = inputStream.read(buffer);
            }
            baos.flush();
            return baos.toByteArray();
        } catch (IOException e) {
			throw new POIException("读取excel流解析时出错",e);
		} finally {
            if (inputStream != null) {
                try {
					inputStream.close();
				} catch (IOException e) {
					System.out.println("警告：关闭读取excel流出错,"+e.getMessage()); 
				}
            }
        }
    }
}
