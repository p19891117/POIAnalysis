package top.tangshitai.excel;

import java.util.Arrays;
import java.util.List;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import top.tangshitai.excel.bean.POIResult;
import top.tangshitai.excel.exception.POIException;

public class POIMainTest{
	private POIMain workBook;
	@Before
	public void before() {
		workBook = new POIMain();
	}
	
	@Test
	public void testAnalysisWorkBook() {
		try {
			List<POIResult> result = workBook.analysisWorkBook(null, "src/test/resources/test.xls",new String[]{"名字","年龄","地址","来源","婚配","结果","说明"},new String[]{"1","2","3"});
			for(POIResult en:result) {
				for(int x=0;x<en.rowLength();x++) {
					System.out.println(Arrays.toString(en.getRow(x)));
				}
				System.out.println("sheet名称："+en.getSheetName());
			}
		}catch (POIException e) {
			e.printStackTrace();
		}
	}

	@Test
	public void testAnalysisWorkBook1() {
		try {
			List<POIResult> result = workBook.analysisWorkBook(null, "src/test/resources/person.xls",new String[]{"姓名","专家级别","专业"});
			for(POIResult en:result) {
				for(int x=0;x<en.rowLength();x++) {
					System.out.println(Arrays.toString(en.getRow(x)));
				}
				System.out.println("sheet名称："+en.getSheetName());
			}
		}catch (POIException e) {
			e.printStackTrace();
		}
	}
	
	@After
	public void after() {
		workBook = null;
	}
}
