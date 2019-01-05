package top.tangshitai.excel;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import top.tangshitai.excel.bean.POIResult;
import top.tangshitai.excel.exception.POIException;
import top.tangshitai.excel.utils.POIUtils;

public class POIMain {
	private SheetProcess sheetProcess = new SheetProcess();
	public List<POIResult> analysisWorkBook(String prefix,String fileName,String[]... sheetConfig) throws POIException {
		try {
			InputStream in = new FileInputStream(POIUtils.path(prefix,fileName));
			ByteArrayInputStream byteIn = new ByteArrayInputStream(POIUtils.readToByte(in));
			Workbook workbook = WorkbookFactory.create(byteIn);
			List<POIResult> result = sheetProcess.analysisSheet(workbook, sheetConfig);
			return result;
		}catch (POIException e) {
			throw e;
		}catch (Exception e) {
			throw new POIException("解析excel["+fileName+"]出错",e);
		}
		
	}
	public List<POIResult> analysisWorkBook(InputStream in,String[]... sheetConfig) throws POIException {
		try {
			ByteArrayInputStream byteIn = new ByteArrayInputStream(POIUtils.readToByte(in));
			Workbook workbook = WorkbookFactory.create(byteIn);
			List<POIResult> result = sheetProcess.analysisSheet(workbook, sheetConfig);
			return result;
		}catch (POIException e) {
			throw e;
		}catch (Exception e) {
			throw new POIException("解析excel出错",e);
		}
		
	}
}
