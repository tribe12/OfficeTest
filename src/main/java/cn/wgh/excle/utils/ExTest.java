package cn.wgh.excle.utils;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.junit.Test;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;

import cn.wgh.Annotation.test1.AnnoProperty;
import cn.wgh.Annotation.test1.ParamMapping;
import cn.wgh.Annotation.test1.User;

public class ExTest {
	@Test
	public void testName() throws Exception {
		Field[] fields = User.class.getDeclaredFields();
		System.out.println(fields);
		List<AnnoProperty> annoPros = new ArrayList<AnnoProperty>();
		for (Field field : fields) {
			ParamMapping exportMapping = field.getAnnotation(ParamMapping.class);
			if (exportMapping == null)
				continue;
			String name = exportMapping.name();
			int sort = exportMapping.sort();
			String property = field.getName();
			System.out.println("------------");
			System.out.println(name);
			System.out.println(sort);
			System.out.println(property);
			AnnoProperty annoPro = new AnnoProperty();
			annoPro.setName(name);
			annoPro.setSort(sort);
			annoPro.setProperty(property);
			annoPros.add(annoPro);
		}
		// 根据sort字段排序
		Collections.sort(annoPros);
		System.out.println(annoPros);
		List<Map<String, String>> columnMapping = new ArrayList<Map<String, String>>();
		if (annoPros != null && !annoPros.isEmpty()) {
			for (AnnoProperty an : annoPros) {
				Map<String, String> mapping = new HashMap<String, String>();
				mapping.put(an.getName(), an.getProperty());
				columnMapping.add(mapping);
			}
		}
		System.out.println("---");
		System.out.println(columnMapping);
	}

	
	
	
	@Test
	public void testName2() throws Exception {
		TitleAndValForTable valTable = new TitleAndValForTable();
		
		List<AnnoProperty> annoPros = getAnnoPropertyForTable(User.class);
		List<User> userDatas = getUserData();
		
		List<String> title = new ArrayList<String>();
		List<List<String>> valueList = new ArrayList<List<String>>();
		JSONArray valJsonArray = new JSONArray();
		if(userDatas!=null && !userDatas.isEmpty()){
			valJsonArray = (JSONArray) JSONArray.toJSON(userDatas);
		}	
		
		for (AnnoProperty an : annoPros) {
			title.add(an.getName());
		}
		JSONObject jsonObj = null;
		for (Object oj : valJsonArray) {
			jsonObj = (JSONObject) oj;
			List<String> val = new ArrayList<String>();
			for (AnnoProperty an : annoPros) {
				val.add(jsonObj.getString(an.getProperty()));
			}
			valueList.add(val);
		}
		
		valTable.setTitle(title);
		valTable.setValueList(valueList);
		
		System.out.println(valTable);
//		return valTable;
	}
	
	
	
	
	
	List<AnnoProperty> getAnnoPropertyForTable(Class<?> clazz) {
		Field[] fields = clazz.getDeclaredFields();
		List<AnnoProperty> annoPros = new ArrayList<AnnoProperty>();
		for (Field field : fields) {
			ParamMapping exportMapping = field.getAnnotation(ParamMapping.class);
			if (exportMapping == null)
				continue;
			String name = exportMapping.name();
			int sort = exportMapping.sort();
			String property = field.getName();
			AnnoProperty annoPro = new AnnoProperty();
			annoPro.setName(name);
			annoPro.setSort(sort);
			annoPro.setProperty(property);
			annoPros.add(annoPro);
		}
		// 根据sort字段排序
		Collections.sort(annoPros);
		return annoPros;
	}

	
	
	List<User> getUserData(){
		User user1 = new User("007", "小A", "7");
		User user2 = new User("008", "小B", "8");
		User user3 = new User("0010", "小C", "7");
		User user4 = new User("002", "小D", "10");
		User user5 = new User("003", "小E", "11");
		User user6 = new User("004", "小F", "7");
		User user7 = new User("005", "小G", "11");
		User user8 = new User("006", "小H", "8");
		User user9 = new User("009", "小J", "9");
		User user10 = new User("001", "小K", "12");
		List<User> users = Arrays.asList(user1, user2, user3, user4
				,user5, user6, user7, user8,user9, user10);
		return users;
	};
	
	
	
	
	
	
	
	
	
	
	@Test
	public void ExportExcelUtilTest() throws Exception {
		//读取类的属性
		List<AnnoProperty> annoPros = ExportExcelUtil.getAnnoPropertyForTable(User.class);
		
		//查询数据
		List<User> userDatas = getUserData();
		JSONArray jsonDatas = (JSONArray) JSONArray.toJSON(userDatas);
		userDatas =null;
		//转换封装title和valueList
		TitleAndValForTable packTableVal = ExportExcelUtil.packTableVal(annoPros, jsonDatas);
		System.out.println(jsonDatas);
		System.out.println(packTableVal);
		//导出
//		HttpServletResponse response = null;
//		ExportExcelUtil.exportExcleFile("文件名", packTableVal.getTitle(),packTableVal.getValueList(), response);
		
	}
	
	
	
	
	
}
