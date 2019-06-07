package cn.wgh.excle.utils;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;

import cn.wgh.Annotation.test1.AnnoProperty;
import cn.wgh.Annotation.test1.ParamMapping;

/**
 * 可用于项目中对列表查询的导出
 * 
 * @author guohui.wang 2017.11
 */
public class ExportExcelUtil {
	/**
	 * 导出excle文件
	 * 
	 * @param fileTitle
	 *            导出的excle文件的名称
	 * @param title
	 *            excle中的列表头
	 * @param valueList
	 *            excle中的列表值
	 * @param response
	 * @throws IOException
	 */
	public static void exportExcleFile(String fileTitle, List<String> title, List<List<String>> valueList,
			HttpServletResponse response) throws IOException {
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet(fileTitle);
		HSSFCellStyle style = wb.createCellStyle();
		sheet.setDefaultColumnWidth(18);// 设置列宽
		style = getCellStyle(wb, style);
		if (title != null && !title.isEmpty()) {
			// 总列数
			int tdLength = title.size();
			// 总行数
			int trLength = 0;
			if (valueList != null && !valueList.isEmpty()) {
				trLength += valueList.size();
				// 创建行
				for (int i = 0; i <= trLength; i++) {
					HSSFRow row = sheet.createRow(i);
					// 创建单元格
					for (int j = 0; j < tdLength; j++) {
						HSSFCell cell = row.createCell(j);
						cell.setCellStyle(style);
						if (i == 0) {
							cell.setCellValue(title.get(j));
						} else {
							cell.setCellValue(valueList.get(i - 1).get(j));
						}
					}
				}
			} else {
				HSSFRow row = sheet.createRow(0);
				for (int j = 0; j < tdLength; j++) {
					HSSFCell cell = row.createCell(j);
					cell.setCellStyle(style);
					cell.setCellValue(title.get(j));
				}
			}
		}

		response.setContentType("application/x-msdownload");
		response.setHeader("Content-Disposition",
				"attachment; filename=" + URLEncoder.encode(fileTitle + ".xls", "UTF-8"));
		OutputStream out = response.getOutputStream();

		out.flush();
		wb.write(out);
		out.close();
	}

	private static HSSFCellStyle getCellStyle(HSSFWorkbook wb, HSSFCellStyle style) {
		HSSFFont font = wb.createFont();
		font.setFontHeightInPoints((short) 11);
		font.setFontName("宋体");
		style.setFont(font);
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style.setBorderBottom((short) 1);
		style.setBorderTop((short) 1);
		style.setBorderLeft((short) 1);
		style.setBorderRight((short) 1);
		return style;
	}

	/**
	 * 读取要导出的类的注解属性
	 * @param clazz
	 * @return
	 */
	public static List<AnnoProperty> getAnnoPropertyForTable(Class<?> clazz) {
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

	
	/**
	 * 封装title和valueList数据
	 * @param annoPros
	 * @param jsonDatas
	 * @return
	 */
	public static TitleAndValForTable packTableVal(List<AnnoProperty> annoPros, JSONArray jsonDatas) {
		TitleAndValForTable valTable = new TitleAndValForTable();
		List<String> title = new ArrayList<String>();
		List<List<String>> valueList = new ArrayList<List<String>>();

		for (AnnoProperty an : annoPros) {
			title.add(an.getName());
		}
		JSONObject jsonObj = null;
		if (jsonDatas != null && !jsonDatas.isEmpty()) {
			for (Object obj : jsonDatas) {
				jsonObj = (JSONObject) obj;
				List<String> val = new ArrayList<String>();
				for (AnnoProperty an : annoPros) {
					val.add(jsonObj.getString(an.getProperty()));
				}
				valueList.add(val);
			}
		}

		valTable.setTitle(title);
		valTable.setValueList(valueList);
		return valTable;
	}

}
