package top.tangshitai.excel;

import java.util.Arrays;
import java.util.List;

import top.tangshitai.excel.bean.POIResult;
import top.tangshitai.excel.exception.POIException;
import top.tangshitai.excel.utils.POIUtils;

public class App {
	public static void main(String[] args) {
		try {
			List<POIResult> result = POIUtils.analysisWorkBook("/home/tst/Desktop/", "test.xls",new String[]{"名字","年龄","地址","来源","婚配","结果","说明"});
			for(POIResult en:result) {
				for(String[] tmp:en.getData()) {
					System.out.println(Arrays.toString(tmp));
				}
				System.out.println("sheet名称："+en.getSheetName()+"-------------------------");
			}
		}catch (POIException e) {
			e.printStackTrace();
		}
	}
	
}