package cn.wgh.excle;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.imageio.ImageIO;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;
import org.junit.Test;

public class 修改模板Excle内容Test {
	String url1 = "E:/test/xlsTest/POITest/ExcleTest/test";
	String url2 = "E:/test/xlsTest/POITest/ExcleTest/修改test" + System.currentTimeMillis() + ".xls";
	@Test
	public void testName() throws Exception {
		
		
			FileInputStream fs = new FileInputStream(url1+"1.xls");// 获取
			POIFSFileSystem ps = new POIFSFileSystem(fs); // 使用POI提供的方法得到excel的信息
			HSSFWorkbook wb = new HSSFWorkbook(ps);
			HSSFSheet sheet = wb.getSheetAt(0); // 获取到工作表，因为一个excel可能有多个工作表
			// HSSFSheet sheet = wb.getSheet("Sheet2"); // 根据名字获取到工作表

			int pLength = sheet.getPhysicalNumberOfRows();
			System.out.println("物理行数 pLength:" + pLength);
			// 总行数
			int trLength = sheet.getLastRowNum();
			System.out.println("总行数：" + trLength);
			trLength = trLength < pLength ? pLength : trLength + 1;
			// 4.得到Excel工作表的行
			Row sheetrow = sheet.getRow(0);
			// 总列数
			int tdLength = sheetrow.getLastCellNum();

			int lastRowNum = sheet.getLastRowNum();
			System.out.println("lastRowNum:" + lastRowNum);
			int addRowCount = 0;//所有数据需增加行数
			int sizeRowCount = 0;//每种数据需要加行数
			int insertRowNum = 0;
			int insertCellNum = 0;
			
			
			int cjInsertRowNum = 0;// 需要插入成绩的行号
			int cjInsertCellNum = 0;// 需要插入成绩的列号
			int pyInsertRowNum = 0;// 需要插入评语的行号
			int pyInsertCellNum = 0;// 需要插入评语的列号

			User user = getUser();
			for (int rowNum = 0; rowNum < pLength; rowNum++) {
				HSSFRow row = sheet.getRow(rowNum);
				int cellLength = row.getLastCellNum();
				for (int cellNum = 0; cellNum < cellLength; cellNum++) {
					HSSFCell cell = row.getCell(cellNum);

					if (cell != null) {
						switch (cell.getStringCellValue()) {
						case "name":
							cell.setCellValue(user.getName());
							break;
						case "sex":
							cell.setCellValue(user.getSex());
							break;
						case "age":
							cell.setCellValue(user.getAge());
							break;
						case "chengji":
							cell.setCellValue(user.getChengji());
							break;
						case "xueyear":
							cjInsertRowNum = rowNum;
							cjInsertCellNum = cellNum;
							break;
						case "pingyu":
							cell.setCellValue(user.getPingyu());
							break;
						default:
							break;
						}
					}
				}
			}
			FileOutputStream out = new FileOutputStream(
					url2);//// 向xls中写数据
			out.flush();
			wb.write(out);
			out.close();
	}
	
	//=================================================================================
	@Test
	public void 删除已存在的Excle的指定列内容_写入新类容1() throws Exception {
		FileInputStream fs = new FileInputStream("C:/Users/dell/Desktop/Base/test1.xls");// 获取
		POIFSFileSystem ps = new POIFSFileSystem(fs); // 使用POI提供的方法得到excel的信息
		HSSFWorkbook wb = new HSSFWorkbook(ps);
		HSSFSheet sheet = wb.getSheetAt(0); // 获取到工作表，因为一个excel可能有多个工作表
		// HSSFSheet sheet = wb.getSheet("Sheet2"); // 根据名字获取到工作表

		int pLength = sheet.getPhysicalNumberOfRows();
		System.out.println("物理行数 pLength:" + pLength);
		// 总行数
		int trLength = sheet.getLastRowNum();
		System.out.println("总行数：" + trLength);
		trLength = trLength < pLength ? pLength : trLength + 1;
		// 4.得到Excel工作表的行
		Row sheetrow = sheet.getRow(0);
		// 总列数
		int tdLength = sheetrow.getLastCellNum();

		int lastRowNum = sheet.getLastRowNum();
		System.out.println("lastRowNum:" + lastRowNum);

		int insertRowNum = 0;// 需要插入的行号
		int insertCellNum = 0;// 需要插入的列号

		User user = getUser();
		for (int rowNum = 0; rowNum < pLength; rowNum++) {
			HSSFRow row = sheet.getRow(rowNum);
			int cellLength = row.getLastCellNum();
			for (int cellNum = 0; cellNum < cellLength; cellNum++) {
				HSSFCell cell = row.getCell(cellNum);

				if (cell != null) {
					switch (cell.getStringCellValue()) {
					case "name":
						cell.setCellValue(user.getName());
						break;
					case "sex":
						cell.setCellValue(user.getSex());
						break;
					case "age":
						cell.setCellValue(user.getAge());
						break;
					case "chengji":
						cell.setCellValue(user.getChengji());
						break;
					case "xueyear":
						insertRowNum = rowNum;
						insertCellNum = cellNum;
						break;
					// case "yuwen":
					// cell.setCellValue(user.getYuwen());
					// break;
					// case "shuxue":
					// cell.setCellValue(user.getShuxue());
					// break;
					// case "yingyu":
					// cell.setCellValue(user.getYingyu());
					// break;
					case "pingyu":
						cell.setCellValue(user.getPingyu());
						break;
					default:
						break;
					}
				}
			}
		}

		System.out.println("insertRowNum:"+insertRowNum);
		System.out.println("insertCellNum:"+insertCellNum);
		
		List<Map<String, Object>> cjList = getChengji();
		int nianji = cjList.size();// 需要插入三条
		

		for (int nj = 1; nj < nianji; nj++) {
			// sheet.shiftRows(insertRowNum, sheet.getLastRowNum(), 1, true,
			// false);
			sheet.shiftRows(insertRowNum, sheet.getLastRowNum(), 1, false, true);
			sheet.createRow(insertRowNum);
		}
		for (int rowNum = insertRowNum; rowNum < insertRowNum + nianji; rowNum++) {
			Map<String, Object> map = cjList.get(rowNum-insertRowNum);
			int cellLength = map.size();
			for (int cellNum = insertCellNum; cellNum < insertCellNum + cellLength; cellNum++) {
				HSSFRow row = sheet.getRow(rowNum);
				HSSFCell cell = row.getCell(cellNum);
				if (cell != null) {
					cell.setCellValue((int)map.get("yuwen"));
				}else{
					HSSFCell cellemp = row.createCell(cellNum);
					cellemp.setCellValue((int)map.get("yuwen"));
				}
			}
		}

		FileOutputStream out = new FileOutputStream(
				"C:/Users/dell/Desktop/Base/test" + System.currentTimeMillis() + ".xls");//// 向xls中写数据
		out.flush();
		wb.write(out);
		out.close();
	}

	
	
	//=================================================================================
	@Test
	public void 删除已存在的Excle的指定列内容_写入新类容2() throws Exception {
		FileInputStream fs = new FileInputStream(url1+"2.xls");// 获取
		POIFSFileSystem ps = new POIFSFileSystem(fs); // 使用POI提供的方法得到excel的信息
		HSSFWorkbook wb = new HSSFWorkbook(ps);
		HSSFSheet sheet = wb.getSheetAt(0); // 获取到工作表，因为一个excel可能有多个工作表
		// HSSFSheet sheet = wb.getSheet("Sheet2"); // 根据名字获取到工作表
		int pLength = sheet.getPhysicalNumberOfRows();
		System.out.println("物理行数 pLength:" + pLength);
		// 总行数
		int trLength = sheet.getLastRowNum();
		System.out.println("总行数：" + trLength);
		trLength = trLength < pLength ? pLength : trLength + 1;
		// 4.得到Excel工作表的行
		Row sheetrow = sheet.getRow(0);
		// 总列数
		int tdLength = sheetrow.getLastCellNum();

		int lastRowNum = sheet.getLastRowNum();
		System.out.println("lastRowNum:" + lastRowNum);

		int cjInsertRowNum = 0;// chengji需要插入的行号
		int cjInsertCellNum = 0;// chengji需要插入的列号
		int pyInsertRowNum = 0;// pingyu需要插入的行号
		int pyInsertCellNum = 0;// pingyu需要插入的列号

		User user = getUser();
		for (int rowNum = 0; rowNum < pLength; rowNum++) {
			HSSFRow row = sheet.getRow(rowNum);
			int cellLength = row.getLastCellNum();
			for (int cellNum = 0; cellNum < cellLength; cellNum++) {
				HSSFCell cell = row.getCell(cellNum);

				if (cell != null) {
					switch (cell.getStringCellValue()) {
					case "name":
						cell.setCellValue(user.getName());
						break;
					case "sex":
						cell.setCellValue(user.getSex());
						break;
					case "age":
						cell.setCellValue(user.getAge());
						break;
					case "chengji":
						cell.setCellValue(user.getChengji());
						break;
					case "xueyear":
						cjInsertRowNum = rowNum;
						cjInsertCellNum = cellNum;
						break;
					case "pingyu":
						pyInsertRowNum = rowNum;
						pyInsertCellNum = cellNum;
						break;
					case "zongjie":
						cell.setCellValue(user.getZongjie());
						break;
					default:
						break;
					}
				}
			}
		}

		//========================插入成绩===========================
		System.out.println("cjInsertRowNum:"+cjInsertRowNum);
		System.out.println("cjInsertCellNum:"+cjInsertCellNum);
		List<Map<String, Object>> cjList = getChengji();
		int cjSize = cjList.size();// 需要插入条数
		//先复制插入行
		for (int nj = 1; nj < cjSize; nj++) {
			// sheet.shiftRows(cjInsertRowNum, sheet.getLastRowNum(), 1, true,
			// false);
			sheet.shiftRows(cjInsertRowNum, sheet.getLastRowNum(), 1, false, true);
			sheet.createRow(cjInsertRowNum);
			HSSFRow row = sheet.getRow(cjInsertRowNum);
			int num = 0;
			Map<String, Object> map = cjList.get(num++);
			int cellLength = map.size();
			for (int cellNum = cjInsertCellNum; cellNum < cjInsertCellNum + cellLength; cellNum++) {
				HSSFCell cell = row.getCell(cellNum);
				if(cell==null){
					cell = row.createCell(cellNum);
				}
				HSSFRow rowemp = sheet.getRow(cjInsertRowNum+1);
				HSSFCell cellemp = rowemp.getCell(cellNum);
				String valueemp = cellemp.getStringCellValue();
				cell.setCellValue(valueemp);
			}
		}
		//再填写数据
		for (int rowNum = cjInsertRowNum; rowNum < cjInsertRowNum + cjSize; rowNum++) {
			Map<String, Object> map = cjList.get(rowNum-cjInsertRowNum);
			int cellLength = map.size();
			for (int cellNum = cjInsertCellNum; cellNum < cjInsertCellNum + cellLength; cellNum++) {
				HSSFRow row = sheet.getRow(rowNum);
				HSSFCell cell = row.getCell(cellNum);
				if (cell == null) {
					cell = row.createCell(cellNum);
//					cell = setCellValue(key);
				}
				String key = cell.getStringCellValue();
				cell.setCellValue((int)map.get(key));
			}
		}
		
		
		
		
		
		
		FileOutputStream out = new FileOutputStream(url2);//// 向xls中写数据
		out.flush();
		wb.write(out);
		out.close();
	}

	
	//=================================================================================
	@Test
	public void 删除已存在的Excle的指定列内容_写入新类容3() throws Exception {
		FileInputStream fs = new FileInputStream(url1+"2.xls");// 获取
		POIFSFileSystem ps = new POIFSFileSystem(fs); // 使用POI提供的方法得到excel的信息
		HSSFWorkbook wb = new HSSFWorkbook(ps);
		HSSFSheet sheet = wb.getSheetAt(0); // 获取到工作表，因为一个excel可能有多个工作表
		// HSSFSheet sheet = wb.getSheet("Sheet2"); // 根据名字获取到工作表
		int pLength = sheet.getPhysicalNumberOfRows();
		System.out.println("物理行数 pLength:" + pLength);
		// 总行数
		int trLength = sheet.getLastRowNum();
		System.out.println("总行数：" + trLength);
		trLength = trLength < pLength ? pLength : trLength + 1;
		// 4.得到Excel工作表的行
		Row sheetrow = sheet.getRow(0);
		// 总列数
		int tdLength = sheetrow.getLastCellNum();

		int lastRowNum = sheet.getLastRowNum();
		System.out.println("lastRowNum:" + lastRowNum);

		int cjInsertRowNum = 0;// chengji需要插入的行号
		int cjInsertCellNum = 0;// chengji需要插入的列号
		int pyInsertRowNum = 0;// pingyu需要插入的行号
		int pyInsertCellNum = 0;// pingyu需要插入的列号

		User user = getUser();
		for (int rowNum = 0; rowNum < pLength; rowNum++) {
			HSSFRow row = sheet.getRow(rowNum);
			int cellLength = row.getLastCellNum();
			for (int cellNum = 0; cellNum < cellLength; cellNum++) {
				HSSFCell cell = row.getCell(cellNum);

				if (cell != null) {
					switch (cell.getStringCellValue()) {
					case "name":
						cell.setCellValue(user.getName());
						break;
					case "sex":
						cell.setCellValue(user.getSex());
						break;
					case "age":
						cell.setCellValue(user.getAge());
						break;
					case "chengji":
						cell.setCellValue(user.getChengji());
						break;
					case "xueyear":
						cjInsertRowNum = rowNum;
						cjInsertCellNum = cellNum;
						break;
					case "pingyu":
						pyInsertRowNum = rowNum;
						pyInsertCellNum = cellNum;
						break;
					case "zongjie":
						cell.setCellValue(user.getZongjie());
						break;
					default:
						break;
					}
				}
			}
		}

		//========================插入成绩===========================
				System.out.println("cjInsertRowNum:"+cjInsertRowNum);
				System.out.println("cjInsertCellNum:"+cjInsertCellNum);
				List<Map<String, Object>> cjList = getChengji();
				int cjSize = cjList.size();// 需要插入条数
				//先复制插入行
				for (int nj = 1; nj < cjSize; nj++) {
					// sheet.shiftRows(cjInsertRowNum, sheet.getLastRowNum(), 1, true,
					// false);
					sheet.shiftRows(cjInsertRowNum, sheet.getLastRowNum(), 1, false, true);
					sheet.createRow(cjInsertRowNum);
					HSSFRow row = sheet.getRow(cjInsertRowNum);
					int num = 0;
					Map<String, Object> map = cjList.get(num++);
					int cellLength = map.size();
					for (int cellNum = cjInsertCellNum; cellNum < cjInsertCellNum + cellLength; cellNum++) {
						HSSFCell cell = row.getCell(cellNum);
						if(cell==null){
							cell = row.createCell(cellNum);
						}
						HSSFRow rowemp = sheet.getRow(cjInsertRowNum+1);
						HSSFCell cellemp = rowemp.getCell(cellNum);
						String valueemp = cellemp.getStringCellValue();
						cell.setCellValue(valueemp);
					}
				}
				//再填写数据
				for (int rowNum = cjInsertRowNum; rowNum < cjInsertRowNum + cjSize; rowNum++) {
					Map<String, Object> map = cjList.get(rowNum-cjInsertRowNum);
					int cellLength = map.size();
					for (int cellNum = cjInsertCellNum; cellNum < cjInsertCellNum + cellLength; cellNum++) {
						HSSFRow row = sheet.getRow(rowNum);
						HSSFCell cell = row.getCell(cellNum);
						if (cell == null) {
							cell = row.createCell(cellNum);
//							cell = setCellValue(key);
						}
						String key = cell.getStringCellValue();
						cell.setCellValue((int)map.get(key));
					}
				}
				
				
		//========================插入评语===========================
		System.out.println("=======插入评语=======pyInsertRowNum:"+pyInsertRowNum);
		System.out.println("=======插入评语=======pyInsertCellNum:"+pyInsertCellNum);
		System.out.println("=======插入评语=======pyInsertRowNum+:"+(pyInsertRowNum+=cjInsertRowNum));
		System.out.println("=======插入评语=======pyInsertCellNum+:"+(pyInsertCellNum+=cjInsertCellNum));

		List<String> pyList = getPingyu();
		int pySize = pyList.size();
//		for (int rowNum = pyInsertRowNum; rowNum < pyInsertRowNum + pySize; rowNum++) {
			sheet.shiftRows(pyInsertRowNum, sheet.getLastRowNum(), 1, false, true);
			sheet.createRow(pyInsertRowNum);
			HSSFRow row = sheet.getRow(pyInsertRowNum);
			int num = 0;
			String  py = pyList.get(num++);
			int cellLength = pyList.size();
//			for (int cellNum = cjInsertCellNum; cellNum < cjInsertCellNum + cellLength; cellNum++) {
				HSSFCell cell = row.getCell(cjInsertCellNum);
				if(cell==null){
					cell = row.createCell(cjInsertCellNum);
				}
				HSSFRow rowemp = sheet.getRow(pyInsertRowNum);
				HSSFCell cellemp = rowemp.getCell(cjInsertCellNum);
				String valueemp = cellemp.getStringCellValue();
				cell.setCellValue(valueemp);
//			}
				
//		}
		
		
		
		
		
		FileOutputStream out = new FileOutputStream(url2);//// 向xls中写数据
		out.flush();
		wb.write(out);
		out.close();
	
}
	
	
	
	
	@Test
	public void 删除已存在的Excle的指定列内容_写入新类容4() throws Exception {
		FileInputStream fs = new FileInputStream(url1+"2.xls");// 获取
		POIFSFileSystem ps = new POIFSFileSystem(fs); // 使用POI提供的方法得到excel的信息
		HSSFWorkbook wb = new HSSFWorkbook(ps);
		HSSFSheet sheet = wb.getSheetAt(0); // 获取到工作表，因为一个excel可能有多个工作表
		// HSSFSheet sheet = wb.getSheet("Sheet2"); // 根据名字获取到工作表
		int pLength = sheet.getPhysicalNumberOfRows();
		System.out.println("物理行数 pLength:" + pLength);
		// 总行数
		int trLength = sheet.getLastRowNum();
		System.out.println("总行数：" + trLength);
		trLength = trLength < pLength ? pLength : trLength + 1;
		// 4.得到Excel工作表的行
		Row sheetrow = sheet.getRow(0);
		// 总列数
		int tdLength = sheetrow.getLastCellNum();

		int lastRowNum = sheet.getLastRowNum();
		System.out.println("lastRowNum:" + lastRowNum);

		int cjInsertRowNum = 0;// chengji需要插入的行号
		int cjInsertCellNum = 0;// chengji需要插入的列号
		int pyInsertRowNum = 0;// pingyu需要插入的行号
		int pyInsertCellNum = 0;// pingyu需要插入的列号

		User user = getUser();
		
		 BufferedImage bufferImg = null;  
		 ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();     
	     bufferImg = ImageIO.read(new File("E:/test/xlsTest/POITest/ExcleTest/pic.jpg"));     
	     ImageIO.write(bufferImg, "jpg", byteArrayOut);  
		 HSSFPatriarch patriarch = sheet.createDrawingPatriarch();     
		
		for (int rowNum = 0; rowNum < pLength; rowNum++) {
			HSSFRow row = sheet.getRow(rowNum);
			int cellLength = row.getLastCellNum();
			for (int cellNum = 0; cellNum < cellLength; cellNum++) {
				HSSFCell cell = row.getCell(cellNum);

				if (cell != null) {
					switch (cell.getStringCellValue()) {
					case "name":
						cell.setCellValue(user.getName());
						break;
					case "sex":
						cell.setCellValue(user.getSex());
						break;
					case "age":
						cell.setCellValue(user.getAge());
						break;
					case "chengji":
						cell.setCellValue(user.getChengji());
						break;
					case "xueyear":
						cjInsertRowNum = rowNum;
						cjInsertCellNum = cellNum;
						break;
					case "pingyu":
						pyInsertRowNum = rowNum;
						pyInsertCellNum = cellNum;
						break;
					case "zongjie":
						cell.setCellValue(user.getZongjie());
						break;
					case "pic":
						
						int rowIndex = cell.getRowIndex();
						
						
//						 HSSFClientAnchor anchor = new HSSFClientAnchor(5, 5, 50, 50,(short) rowIndex, cellNum, (short) 5, 8);     
//						 HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0,0,(short) cellNum, rowIndex, (short) 5, 8);     
						 HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0,0,(short) cellNum, rowIndex, (short)(cellNum+1), rowIndex+3);     
				            //插入图片    
				         patriarch.createPicture(anchor, wb.addPicture(byteArrayOut.toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));   
						
						
//						cell.setCellValue(user.getZongjie());
						break;
					default:
						break;
					}
				}
			}
		}

		//========================插入成绩===========================
				System.out.println("cjInsertRowNum:"+cjInsertRowNum);
				System.out.println("cjInsertCellNum:"+cjInsertCellNum);
				List<Map<String, Object>> cjList = getChengji();
				int cjSize = cjList.size();// 需要插入条数
				//先复制插入行
				for (int nj = 1; nj < cjSize; nj++) {
					// sheet.shiftRows(cjInsertRowNum, sheet.getLastRowNum(), 1, true,
					// false);
					sheet.shiftRows(cjInsertRowNum, sheet.getLastRowNum(), 1, false, true);
					sheet.createRow(cjInsertRowNum);
					HSSFRow row = sheet.getRow(cjInsertRowNum);
					int num = 0;
					Map<String, Object> map = cjList.get(num++);
					int cellLength = map.size();
					for (int cellNum = cjInsertCellNum; cellNum < cjInsertCellNum + cellLength; cellNum++) {
						HSSFCell cell = row.getCell(cellNum);
						if(cell==null){
							cell = row.createCell(cellNum);
						}
						HSSFRow rowemp = sheet.getRow(cjInsertRowNum+1);
						HSSFCell cellemp = rowemp.getCell(cellNum);
						String valueemp = cellemp.getStringCellValue();
						cell.setCellValue(valueemp);
					}
				}
				//再填写数据
				for (int rowNum = cjInsertRowNum; rowNum < cjInsertRowNum + cjSize; rowNum++) {
					Map<String, Object> map = cjList.get(rowNum-cjInsertRowNum);
					int cellLength = map.size();
					for (int cellNum = cjInsertCellNum; cellNum < cjInsertCellNum + cellLength; cellNum++) {
						HSSFRow row = sheet.getRow(rowNum);
						HSSFCell cell = row.getCell(cellNum);
						if (cell == null) {
							cell = row.createCell(cellNum);
//							cell = setCellValue(key);
						}
						String key = cell.getStringCellValue();
						cell.setCellValue((int)map.get(key));
					}
				}
				
				
		//========================插入评语===========================
		System.out.println("=======插入评语=======pyInsertRowNum:"+pyInsertRowNum);
		System.out.println("=======插入评语=======pyInsertCellNum:"+pyInsertCellNum);
		System.out.println("=======插入评语=======pyInsertRowNum+:"+(pyInsertRowNum+=cjInsertRowNum));
		System.out.println("=======插入评语=======pyInsertCellNum+:"+(pyInsertCellNum+=cjInsertCellNum));

		List<String> pyList = getPingyu();
		int pySize = pyList.size();
//		for (int rowNum = pyInsertRowNum; rowNum < pyInsertRowNum + pySize; rowNum++) {
			sheet.shiftRows(pyInsertRowNum, sheet.getLastRowNum(), 1, false, true);
			sheet.createRow(pyInsertRowNum);
			HSSFRow row = sheet.getRow(pyInsertRowNum);
			int num = 0;
			String  py = pyList.get(num++);
			int cellLength = pyList.size();
//			for (int cellNum = cjInsertCellNum; cellNum < cjInsertCellNum + cellLength; cellNum++) {
				HSSFCell cell = row.getCell(cjInsertCellNum);
				if(cell==null){
					cell = row.createCell(cjInsertCellNum);
				}
				HSSFRow rowemp = sheet.getRow(pyInsertRowNum);
				HSSFCell cellemp = rowemp.getCell(cjInsertCellNum);
				String valueemp = cellemp.getStringCellValue();
				cell.setCellValue(valueemp);
//			}
				
//		}
		
		
		
		
		
		FileOutputStream out = new FileOutputStream(url2);//// 向xls中写数据
		out.flush();
		wb.write(out);
		out.close();
	
}
	

	private void addRow(HSSFSheet sheet, int tdLength, int insertRowNum, int insertCellNum, int sizeRowCount) {
		for (int addNum = 0; addNum < sizeRowCount; addNum++) {
			sheet.shiftRows(insertRowNum, sheet.getLastRowNum(), 1, false, true);
			sheet.createRow(insertRowNum);
			HSSFRow row = sheet.getRow(insertRowNum);
			int cellLength = tdLength;
			for (int cellNum = insertCellNum; cellNum < insertCellNum + cellLength; cellNum++) {
				HSSFCell cell = row.getCell(cellNum);
				if (cell == null) {
					cell = row.createCell(cellNum);
				}
				HSSFRow rowemp = sheet.getRow(insertRowNum + 1);
				HSSFCell cellemp = rowemp.getCell(cellNum);
				String valueemp = cellemp.getStringCellValue();
				cell.setCellValue(valueemp);
			}
		}
	}
	
	
	private List<Map<String, Object>> getChengji() {
		List<Map<String, Object>> cjList = new ArrayList<Map<String, Object>>();
		for (int i = 0; i < 3; i++) {
			Map<String, Object> cj = new HashMap<String, Object>();
			cj.put("xueyear", 2015+i);
			cj.put("yuwen", 80 + i);
			cj.put("shuxue", 90 + i);
			cj.put("yingyu", 60 + i);
			cjList.add(cj);
		}
		return cjList;
	}

	private User getUser() {
		User user = new User("001", "小明", "男", "18", "2016", "85", "93", "62", "我的自我总结内容。。。");
		return user;
	}

	
	private List<String> getPingyu() {
		List<String> pyList = new ArrayList<String>();
		for (int i = 0; i < 3; i++) {
			pyList.add("注意加强英语学习");
			pyList.add("加油啊");
			pyList.add("查漏补缺");
		}
		return pyList;
	}
	
}
