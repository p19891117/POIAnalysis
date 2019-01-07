package top.tangshitai.excel;

import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import top.tangshitai.excel.bean.POIResult;
import top.tangshitai.excel.exception.POIException;

public class RowProcess {
	private CellProcess cellProcess = new CellProcess();
	public POIResult analysisRow(String[] headParams,Sheet sheet,FormulaEvaluator evaluator) throws POIException {
		if(headParams==null||headParams.length<=0) throw new POIException("excel的sheet["+sheet.getSheetName()+"]head没有配置");
		int cellLength = headParams.length;//col长度
		POIResult poiResult = new POIResult(sheet.getSheetName(),cellLength,headParams);
		for(int rowIndex=0;rowIndex<sheet.getLastRowNum();rowIndex++) {
			Row row = sheet.getRow(rowIndex);
			String[] rowResult = cellProcess.analysisCell(cellLength, row, evaluator);
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
}
