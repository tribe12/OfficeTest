package cn.wgh.Annotation;

import java.util.ArrayList;
import java.util.List;

import org.junit.Test;

import cn.wgh.Annotation.test1.AnnotationQuery;
import cn.wgh.Annotation.test1.User;

public class AnnotationTest {
	@Test
	public void testName() throws Exception {
		List<String> keyList = getKeyList();
		List<List<String>> valueList = getValueList();
		int valueSize = valueList.size();

		if(keyList!=null && !keyList.isEmpty()){
			// 总列数
			int tdLength = keyList.size();
			// 总行数
			int trLength = 0;
			trLength += valueList.size();
			// 创建行
			for (int i = 0; i < trLength; i++) {
				AnnotationQuery<User> annotationQuery = new AnnotationQuery<User>();
				// 创建单元格
				for (int j = 0; j < tdLength; j++) {
						String key = keyList.get(j);
						String value = valueList.get(i).get(j);
						AnnotationQuery<User> setColumnMapping = annotationQuery.setColumnMapping(key,value);
					}
				}
			}
	}

	private List<String> getKeyList() {
		List<String> key = new ArrayList<String>();
		key.add("姓名");
		key.add("年龄");
		return key;
	}

	private List<List<String>> getValueList() {
		List<List<String>> list = new ArrayList<List<String>>();

		List<String> value = new ArrayList<String>();
		value.add("小明");
		value.add("17");

		list.add(value);

		return list;
	}

}
