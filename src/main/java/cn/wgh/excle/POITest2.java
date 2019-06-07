package cn.wgh.excle;

import static org.junit.Assert.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.junit.Test;

public class POITest2 {
	@Test
	public void writeExistExcel() throws Exception {
		FileInputStream fs = new FileInputStream("E:/test/xlsTest/XSSF/students_CreateExcel.xls");// 获取
//		FileInputStream fs = new FileInputStream("C:\\Users\\Lenovo\\Desktop\\ccccccccccccccccccccccccc\\工资表 (2222.8位数超长)10条.xlsx");// 获取
//		String fileUrl = "C:\\Users\\Lenovo\\Desktop\\ccccccccccccccccccccccccc\\工资表 (2222.8位数超长)10条.xlsx";
		
		
		POIFSFileSystem ps = new POIFSFileSystem(fs); // 使用POI提供的方法得到excel的信息
		HSSFWorkbook wb = new HSSFWorkbook(ps);
		HSSFSheet sheet = wb.getSheetAt(0); // 获取到工作表，因为一个excel可能有多个工作表
//		HSSFSheet sheet = wb.getSheet("Sheet2"); // 根据名字获取到工作表
		HSSFRow row = sheet.getRow(0); // 获取第一行（excel中的行默认从0开始，所以这就是为什么，一个excel必须有字段列头），即，字段列头，便于赋值
		// 分别得到最后一行的行号，和一条记录的最后一个单元格
		System.out.println("sheet.getLastRowNum():"+sheet.getLastRowNum());
		System.out.println("row.getLastCellNum():"+row.getLastCellNum());

		FileOutputStream out = new FileOutputStream("E:/test/xlsTest/XSSF/students_CreateExcel.xls"); // 向xls中写数据
		row = sheet.createRow((short) (sheet.getLastRowNum() + 1)); // 在现有行号后追加数据
		row.createCell(0).setCellValue("旺财"); // 设置第一个（从0开始）单元格的数据
		row.createCell(1).setCellValue(24); // 设置第二个（从0开始）单元格的数据
		row.createCell(2).setCellValue("啦啦啦！"); // 设置第三个（从0开始）单元格的数据

		out.flush();
		wb.write(out);
		out.close();
		System.out.println("row.getPhysicalNumberOfCells():"+row.getPhysicalNumberOfCells());
		System.out.println("row.getLastCellNum():"+row.getLastCellNum());
	}
	
	@Test
	public void 删除指定行() throws Exception {
		/**
		 * 一般情况下，删除行时会面临两种情况：删除行内容但保留行位置、整行删除（删除后下方单元格上移）。
		 * 对应的删除方法分别是：removeRow()及shiftRow(startRow,endRow,shiftCount)
		 */
		try {
			FileInputStream is = new FileInputStream("E:/test/xlsTest/XSSF/students_CreateExcel.xls");
			HSSFWorkbook workbook = new HSSFWorkbook(is);
			HSSFSheet sheet = workbook.getSheetAt(0);
//			sheet.removeRow(sheet.getRow(1));//删除第2行
			sheet.shiftRows(1, sheet.getLastRowNum(), -1); // 删除第2行
			FileOutputStream os = new FileOutputStream("E:/test/xlsTest/XSSF/students_CreateExcel.xls");
			workbook.write(os);
			is.close();
			os.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	
	
	@Test
	public void 删除指定列内容_保留空值() throws Exception {
		/**
		 */
		try {
			FileInputStream is = new FileInputStream("E:/test/xlsTest/XSSF/students_CreateExcel.xls");
			HSSFWorkbook workbook = new HSSFWorkbook(is);
			HSSFSheet sheet = workbook.getSheetAt(0);
			int lastRowNum = sheet.getLastRowNum();
			System.out.println("lastRowNum:"+lastRowNum);
			//指定要删除的列
			int cellnum = 1;
			for (int i = 0; i <= lastRowNum; i++) {
				HSSFRow row = sheet.getRow(i);
				HSSFCell cell = row.getCell(cellnum);
				if (cell != null) {//判断cell为空值:if(cell==null||cell.equals("")||cell.getCellType() ==HSSFCell.CELL_TYPE_BLANK)
					row.removeCell(cell);
				}
			}
			FileOutputStream os = new FileOutputStream("E:/test/xlsTest/XSSF/students_CreateExcel.xls");
			workbook.write(os);
			is.close();
			os.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	
	
	@Test
	public void 删除指定列_后面左移动() throws Exception {
		/**
		 */
		try {
			FileInputStream is = new FileInputStream("E:/test/xlsTest/XSSF/students_CreateExcel.xls");
			HSSFWorkbook workbook = new HSSFWorkbook(is);
			HSSFSheet sheet = workbook.getSheetAt(0);
			int lastRowNum = sheet.getLastRowNum();
			System.out.println("lastRowNum:"+lastRowNum);
			//指定要删除的列

			
			
			
			
			FileOutputStream os = new FileOutputStream("E:/test/xlsTest/XSSF/students_CreateExcel.xls");
			workbook.write(os);
			is.close();
			os.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	
	
	@Test
	public void writeExistExcel2() throws Exception {
		FileInputStream fs = new FileInputStream("E:/test/xlsTest/XSSF/students_CreateExcel.xls");// 获取
		POIFSFileSystem ps = new POIFSFileSystem(fs); // 使用POI提供的方法得到excel的信息
		HSSFWorkbook wb = new HSSFWorkbook(ps);
		HSSFSheet sheet = wb.getSheetAt(0); // 获取到工作表，因为一个excel可能有多个工作表
//		HSSFSheet sheet = wb.getSheet("Sheet2"); // 根据名字获取到工作表
		HSSFRow row = sheet.getRow(0); // 获取第一行（excel中的行默认从0开始，所以这就是为什么，一个excel必须有字段列头），即，字段列头，便于赋值
		
//		sheet.get
		
		// 分别得到最后一行的行号，和一条记录的最后一个单元格
		System.out.println("sheet.getLastRowNum():"+sheet.getLastRowNum());
//		System.out.println("row.getLastCellNum():"+row.getLastCellNum());

		FileOutputStream out = new FileOutputStream("E:/test/xlsTest/XSSF/students_CreateExcel.xls"); // 向xls中写数据
		row = sheet.createRow((short) (sheet.getLastRowNum() + 1)); // 在现有行号后追加数据
		row.createCell(0).setCellValue("旺财"); // 设置第一个（从0开始）单元格的数据
		row.createCell(1).setCellValue(24); // 设置第二个（从0开始）单元格的数据
		row.createCell(2).setCellValue("啦啦啦！"); // 设置第三个（从0开始）单元格的数据

		out.flush();
		wb.write(out);
		out.close();
		System.out.println("row.getPhysicalNumberOfCells():"+row.getPhysicalNumberOfCells());
		System.out.println("row.getLastCellNum():"+row.getLastCellNum());
	}
	
	
	
}
