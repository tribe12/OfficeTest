package cn.wgh.excle;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;

public class EasyExcleTest {
	@Test
	public void writeExcle() throws Exception {
		OutputStream out = new FileOutputStream("F:\\test\\office\\excle\\EasyTestOut.xlsx");
		ExcelWriter excelWriter = EasyExcelFactory.getWriter(out);
		Sheet sheet1 = new Sheet(1, 0, WriteModel.class);

		sheet1.setSheetName("第一个sheet");

		// 写入数据到excelWriter上下文中
		List<WriteModel> data = creatModelList();
		excelWriter.write(data, sheet1);

		excelWriter.finish();
		out.close();
	}

	/**
	 * 动态生成
	 * 
	 * @throws Exception
	 */
	@Test
	public void writeExcle2() throws Exception {
		OutputStream out = new FileOutputStream("F:\\test\\office\\excle\\EasyTestOut2.xlsx");
		ExcelWriter excelWriter = EasyExcelFactory.getWriter(out);
		Sheet sheet1 = new Sheet(1, 0);
		sheet1.setSheetName("第一个sheet");
		Table table1 = new Table(1);
		table1.setTableStyle(DataUtil.creatTableStyle());
		table1.setHead(DataUtil.createTestListStringHead());
		excelWriter.write1(createDynamicModelList(), sheet1, table1);
		excelWriter.merge(5, 7, 0, 4);
		excelWriter.finish();
		out.close();
	}

	/**
	 * 创建数据
	 */
	private List<WriteModel> creatModelList() {
		List<WriteModel> writeModels = new ArrayList<WriteModel>();
		writeModels.add(WriteModel.builder().name("小明").age("10").password("xm123").build());
		writeModels.add(WriteModel.builder().name("小强").age("11").password("q54321").build());
		writeModels.add(WriteModel.builder().name("旺财").age("8").password("asjfbj").build());
		writeModels.add(WriteModel.builder().name("狗蛋").age("9").password("sahjdfdj").build());

		return writeModels;
	}

	/**
	 * 无注解的实体类
	 *
	 * @return
	 */
	private List<List<Object>> createDynamicModelList() {
		// 所有行数据
		List<List<Object>> rows = new ArrayList<List<Object>>();

		for (int i = 0; i < 30; i++) {
			// 一行数据
			List<Object> row = new ArrayList<Object>();
			row.add("字符串" + i);
			row.add("衣服" + i);
			row.add("裤子" + i);
			row.add("沙发" + i);
			row.add("自行车" + i);
			rows.add(row);
		}

		return rows;
	}

}
