package cn.wgh.excle;

import static org.junit.Assert.*;

import java.util.ArrayList;

import org.junit.Test;

public class ExportExcelTest {
	@Test
	public void testName() throws Exception {
		
		String[] headers = {"标题一","标题二","标题三",};
		String[] cols = {"一","二","三",};
		
//		ExportExcel.exportExcel("testExcle", headers, cols, dataset, constant, out);
		ExportExcel.exportExcel("testExcle", headers, cols, null, null, null);
	}
}
