package top.tangshitai.excel;

import java.math.BigInteger;
import java.text.DateFormat;
import java.util.Date;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

import top.tangshitai.excel.exception.POIException;

public class CellProcess {
	public String[] analysisCell(int cellLength,Row row,FormulaEvaluator evaluator) throws POIException {
		String[] cellsContent = new String[cellLength];
		for(int cellIndex=0;cellIndex<cellLength;cellIndex++) {
			Cell cell = row.getCell(cellIndex);
			String cellResult = getCellStrValue(cell,evaluator);
			cellsContent[cellIndex]=StringUtils.isBlank(cellResult)?"":cellResult.trim();
		}
		return cellsContent;
	}
	public String getCellStrValue(Cell cell,FormulaEvaluator evaluator) throws POIException {
		if(cell==null) throw new POIException("poi的Cell对象不能为空");
		if(evaluator==null) throw new POIException("poi的FormulaEvaluator对象不能为空");
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
}
