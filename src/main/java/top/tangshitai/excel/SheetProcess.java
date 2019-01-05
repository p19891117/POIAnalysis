package top.tangshitai.excel;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import top.tangshitai.excel.bean.POIResult;
import top.tangshitai.excel.exception.POIException;

public class SheetProcess {
	private RowProcess rowProcess = new RowProcess();
	public List<POIResult> analysisSheet(Workbook workbook ,String[][] sheetCfg) throws POIException {
		List<POIResult> poirs = new ArrayList<>();
		if(sheetCfg==null||sheetCfg.length<1) new POIException("excel的每个sheet的head没有配置");
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		for(int sheetIndex=0;sheetIndex<workbook.getNumberOfSheets();sheetIndex++) {
			Sheet sheet = workbook.getSheetAt(sheetIndex);
			if(sheetCfg.length==1) {
				poirs.add(rowProcess.analysisRow(sheetCfg[0], sheet, evaluator));
			}else {
				poirs.add(rowProcess.analysisRow(sheetCfg[sheetIndex], sheet, evaluator));
			}
		}
		return poirs;
	}
}
