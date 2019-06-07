package cn.wgh.word;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.net.URLConnection;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.junit.Test;

public class DocumentHandlerTest {
	
	@Test
	public void 根据模板导出Word() throws Exception {
		Map<String, Object> dataMap = new HashMap<String, Object>();
		
		dataMap.put("userName", "旺财11");
		dataMap.put("age", "113");
		dataMap.put("phoneNo", "15212345678111");
		dataMap.put("groupName", "技术部11");
		dataMap.put("detail", "啦啦啦111");
		
		String imageUrl = "E:\\test\\wordTest\\雷神.jpg";
		
//		// 方法一、将网络图片导入wolrd
//		URL url = new URL(imageUrl);
//		// 打开网络输入流
//		URLConnection conn = url.openConnection();
//		// 设置超时间为3秒
//		conn.setConnectTimeout(3 * 1000);
//		// 防止屏蔽程序抓取而返回403错误
//		conn.setRequestProperty("User-Agent", "Mozilla/4.0 (compatible; MSIE 5.0; Windows NT; DigExt)");
//		// 得到输入流
//		InputStream inputStream = conn.getInputStream();
//		// 获取自己数组
//		byte[] data = readInputStream(inputStream); 
        
        //方法二、将本地图片导入wolrd，打开本地输入流
        FileInputStream in = new FileInputStream(imageUrl);
        byte[] data = new byte[in.available()];
        in.read(data);
        in.close();
		
		
		
		
		dataMap.put("detail", in);
		
		
		
		
		
		DocumentHandler documentHandler = new DocumentHandler();
		documentHandler.createDoc("ftlTemp","带图片模板.ftl",dataMap,  "E:\\test\\wordTest\\outFile_"+(new Date().getTime())+".doc");
	}
	
	
	
	
	
	/**
     * 
     * @Title: readInputStream 
     * @Description: 将网络图片流转换成数组 
     * @param @param inputStream
     * @param @return
     * @param @throws IOException
     * @return byte[]
     * @throws
     */
    public static  byte[] readInputStream(InputStream inputStream) throws IOException {    
        byte[] buffer = new byte[1024];    
        int len = 0;    
        ByteArrayOutputStream bos = new ByteArrayOutputStream();    
        while((len = inputStream.read(buffer)) != -1) {    
            bos.write(buffer, 0, len);    
        }    
        bos.close();    
        return bos.toByteArray();    
    }    
	
	
	
}
