package cn.wgh.word;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.HashMap;
import java.util.Map;

import org.junit.Test;

import freemarker.template.Configuration;
import freemarker.template.DefaultObjectWrapper;
import freemarker.template.Template;
import freemarker.template.TemplateExceptionHandler;

//目前没成功



public class FreemarkerWordTest {
	
//	private Configuration configuration = null;
	private Configuration configuration =  new Configuration();
	
	public static void DocUtil() {
		Configuration configuration = new Configuration();
		configuration.setDefaultEncoding("utf-8");
	}
	
	@Test
	public void testName() throws Exception {
		String downloadType = "导出wordTest.xml";
		String savePath = "E:\\test\\wordTest\\";
		
		Map<String, Object> dataMap = new HashMap<String, Object>();
		
		dataMap.put("userName", "旺财");
		dataMap.put("age", "3");
		dataMap.put("phoneNo", "15212345678");
		dataMap.put("groupName", "技术部");
		dataMap.put("detail", "啦啦啦");
		
		Template template = null;
		
		configuration.setClassForTemplateLoading(this.getClass(), "/word/templete");
		configuration.setObjectWrapper(new DefaultObjectWrapper());
		configuration.setTemplateExceptionHandler(TemplateExceptionHandler.IGNORE_HANDLER);
		template = configuration.getTemplate(downloadType);
		
		File file = new File(savePath);
		Writer out = null;
		
		out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(file),"utf-8"));
		template.process(dataMap, out);
		file.delete();
	}
	
	
	
}
