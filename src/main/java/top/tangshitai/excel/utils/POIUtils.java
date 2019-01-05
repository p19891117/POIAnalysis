package top.tangshitai.excel.utils;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import org.apache.commons.lang3.StringUtils;

import top.tangshitai.excel.App;
import top.tangshitai.excel.exception.POIException;

public class POIUtils {
	public static String path(String prefix,String filename) throws POIException {
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
	public static byte[] readToByte(InputStream inputStream) throws POIException{
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
