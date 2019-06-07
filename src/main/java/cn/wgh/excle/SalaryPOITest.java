package cn.wgh.excle;

import static org.junit.Assert.*;

import java.util.List;

import org.junit.Test;

public class SalaryPOITest {
	@Test
	public void testName() throws Exception {
		String fileUrl = "C:\\Users\\Lenovo\\Desktop\\33333333333333.xlsx";
		
		
//		List<List<String>> excelAsFile = SalaryPOIUtils.getExcelAsFile0823解决公式及小数点问题(fileUrl,true,true);
//		List<List<String>> excelAsFile = SalaryPOIUtils.getExcelAsFile0816解决不同数据转换问题(fileUrl,true,true);
//		List<List<String>> excelAsFile = SalaryPOIUtils.getExcelAsFile0731(fileUrl,true,true);
		List<List<String>> excelAsFile = SalaryPOIUtils.getExcelAsFile0915解决空白行问题(fileUrl,true,true);
		
		for (List<String> list : excelAsFile) {
			System.err.println(list);
		}
		
	}
}
