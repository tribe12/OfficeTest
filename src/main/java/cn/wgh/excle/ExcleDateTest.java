package cn.wgh.excle;

import java.util.List;

import org.junit.Test;

public class ExcleDateTest {
	@Test
	public void testName() throws Exception {
		String fileUrl ="C:\\Users\\Lenovo\\Desktop\\Base\\日期.xls";
//		fileUrl =	"C:/Users/Lenovo/Desktop/gzt/导入test.xls";
		
		List<List<String>> excelAsFile = POIUtils.getExcelAsFile(fileUrl);
		for (List<String> list : excelAsFile) {
			System.out.println(list);
		}
	}
}
