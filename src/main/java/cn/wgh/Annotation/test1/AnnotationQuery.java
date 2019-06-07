package cn.wgh.Annotation.test1;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class AnnotationQuery<T> {

	private List<Map<String, String>> columnMapping = new ArrayList<Map<String, String>>();
	
	private List<AnnoProperty> colums = new ArrayList<AnnoProperty>();
	
	
	public AnnotationQuery<T> setColumnMapping(String text,String property){
		Map<String, String> mapping = new HashMap<String, String>();
		mapping.put(text, property);
		this.columnMapping.add(mapping);
		return this;
	}
	
	public AnnotationQuery<T> useAnnotation(Class<T> clazz){
		//从对象注解中去获取
		Field [] fields = clazz.getDeclaredFields();
		for(Field field : fields){
			ParamMapping paramMapping = field.getAnnotation(ParamMapping.class);
			if(paramMapping != null){
				String name = paramMapping.name();
				int sort = paramMapping.sort();
				String property = field.getName();
				AnnoProperty colum = new AnnoProperty();
				colum.setName(name);
				colum.setSort(sort);
				colum.setProperty(property);
				this.colums.add(colum);
			}
		}
		
		Collections.sort(this.colums);
		for(AnnoProperty colum : this.colums){
			this.setColumnMapping(colum.getName(), colum.getProperty());
		}
		return this;
	}
	


	

}
