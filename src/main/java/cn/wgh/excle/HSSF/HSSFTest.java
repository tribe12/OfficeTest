package cn.wgh.excle.HSSF;

import static org.junit.Assert.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.junit.Test;

public class HSSFTest {
	
	/**
	 * 得到Excel，并解析内容
	 * 
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	@SuppressWarnings("resource")
	@Test
	public void getExcelAsFile() throws FileNotFoundException, IOException {
		// 1.得到Excel常用对象
		// POIFSFileSystem fs = new POIFSFileSystem(new
		// FileInputStream("d:/FTP/test.xls"));
//		FileInputStream fileInputStream = new FileInputStream("E:/test/xlsTest/HSSF/stuTest2.xls");
		FileInputStream fileInputStream = new FileInputStream("E:/test/xlsTest/HSSF/sxlt.xls");
		POIFSFileSystem fs = new POIFSFileSystem(fileInputStream);
		// 2.得到Excel工作簿对象
		HSSFWorkbook wb = new HSSFWorkbook(fs);
		// 3.得到Excel工作表对象
		HSSFSheet sheet = wb.getSheetAt(0);
		//物理行数
		int pLength = sheet.getPhysicalNumberOfRows();
		System.out.println("物理行数 pLength:"+pLength);
		// 总行数
		int trLength = sheet.getLastRowNum();
		System.out.println("总行数:"+trLength);
		// 4.得到Excel工作表的行
		HSSFRow row = sheet.getRow(0);
		// 总列数
		int tdLength = row.getLastCellNum();
		
		System.out.println("总列数:"+tdLength);
		
		// 5.得到Excel工作表指定行的单元格
		HSSFCell cell = row.getCell((short) 1);
		// 6.得到单元格样式
		CellStyle cellStyle = cell.getCellStyle();
		
		System.out.println("------------------------------------------------------------------------");
		for (int i = 0; i < pLength; i++) {
			// 得到Excel工作表的行
			HSSFRow row1 = sheet.getRow(i);
			for (int j = 0; j < tdLength; j++) {

				// 得到Excel工作表指定行的单元格
				HSSFCell cell1 = row1.getCell(j);

				/**
				 * 为了处理：Excel异常Cannot get a text value from a numeric cell
				 * 将所有列中的内容都设置成String类型格式
				 */
				if (cell1 != null) {
					cell1.setCellType(Cell.CELL_TYPE_STRING);
					// 获得每一列中的值
					System.out.print(cell1.getStringCellValue() + "\t\t");
				}else{
					System.out.print("\t\t");
				}

			}
			System.out.println();
		}
	}
	
	
	
	
	

	/**
	 * 创建Excel，并写入内容
	 */
	@Test
	public void CreateExcel() throws Exception {
		// 1.创建Excel工作薄对象
		HSSFWorkbook wb = new HSSFWorkbook();
		// 2.创建Excel工作表对象
		HSSFSheet sheet = wb.createSheet("new Sheet");
		// 3.创建Excel工作表的行
		HSSFRow row = sheet.createRow(6);
		// 4.创建单元格样式
		CellStyle cellStyle = wb.createCellStyle();
		// 设置这些样式
		cellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

		// 5.创建Excel工作表指定行的单元格
		row.createCell(0).setCellStyle(cellStyle);
		// 6.设置Excel工作表的值
		row.createCell(0).setCellValue("插点值aaaa");

		row.createCell(1).setCellStyle(cellStyle);
		row.createCell(1).setCellValue("插点值bbbb");

		// 设置sheet名称和单元格内容
		wb.setSheetName(0, "第一张工作表");
		// 设置单元格内容 cell.setCellValue("单元格内容");

		// 最后一步，将文件存到指定位置
		try {
			FileOutputStream fout = new FileOutputStream("E:/test/xlsTest/HSSF/students_CreateExcel.xls");
			wb.write(fout);
			fout.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * 创建Excel的实例
	 * 
	 * @throws ParseException
	 */
	@Test
	public void CreateExcelDemo1() throws ParseException {
		List list = new ArrayList();
		SimpleDateFormat df = new SimpleDateFormat("yyyy-mm-dd");
		Student user1 = new Student(1, "张三", 16, true, df.parse("1997-03-12"));
		Student user2 = new Student(2, "李四", 17, true, df.parse("1996-08-12"));
		Student user3 = new Student(3, "王五", 26, false, df.parse("1985-11-12"));
		list.add(user1);
		list.add(user2);
		list.add(user3);

		// 第一步，创建一个webbook，对应一个Excel文件
		HSSFWorkbook wb = new HSSFWorkbook();
		// 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
		HSSFSheet sheet = wb.createSheet("学生表一");
		// 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short
		sheet.setDefaultColumnWidth(18);//设置列宽
//		sheet.setDefaultRowHeight((short)400);
		HSSFRow row = sheet.createRow((int) 0);
		
		// 第四步，创建单元格，并设置值表头 设置表头居中
		HSSFCellStyle style = wb.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式

		HSSFCell cell = row.createCell((short) 0);
		cell.setCellValue("学号");
		cell.setCellStyle(style);
		cell = row.createCell((short) 1);
		cell.setCellValue("姓名");
		cell.setCellStyle(style);
		cell = row.createCell((short) 2);
		cell.setCellValue("年龄");
		cell.setCellStyle(style);
		cell = row.createCell((short) 3);
		cell.setCellValue("性别");
		cell.setCellStyle(style);
		cell = row.createCell((short) 4);
		cell.setCellValue("生日");
		cell.setCellStyle(style);

		
		HSSFCellStyle style2 = wb.createCellStyle();
		style2 = getCellStyle(wb, style2);
		
		// 第五步，写入实体数据 实际应用中这些数据从数据库得到，

		for (int i = 0; i < list.size(); i++) {
			row = sheet.createRow((int) i + 1);
			Student stu = (Student) list.get(i);
			// 第四步，创建单元格，并设置值
			row.createCell((short) 0).setCellValue((double) stu.getId());
			row.createCell((short) 1).setCellValue(stu.getName());
			row.createCell((short) 2).setCellValue((double) stu.getAge());
			row.createCell((short) 3).setCellValue(stu.isSex() == true ? "男" : "女");
			cell = row.createCell((short) 4);
//			cell.setCellValue(new SimpleDateFormat("yyyy-mm-dd").format(stu.getBirthday()));
			cell.setCellValue("hasjjjjjjjjj级记者吧处罚级别2");
			cell.setCellStyle(style2);
		}
		// 第六步，将文件存到指定位置
		try {
			Date date = new Date();
			long time = date.getTime();
			FileOutputStream fout = new FileOutputStream("E:/test/xlsTest/HSSF/students_CreateExcelDemo"+time+".xls");
			wb.write(fout);
			fout.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}






	private HSSFCellStyle getCellStyle(HSSFWorkbook wb,HSSFCellStyle style) {
		HSSFFont font = wb.createFont();
		
		
		font.setFontHeightInPoints((short) 12);
        font.setFontName("仿宋_GB2312");
//		font.setFontName("宋体");

		style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式
		style.setFont(font);
		
		
		
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
//		style.setBorderBottom((short) 1);  
//		style.setBorderTop((short) 1);  
//		style.setBorderLeft((short) 1);  
//		style.setBorderRight((short) 1);  
		
		
//		style.setWrapText(true);//根据内容自动换行
		return style;
	}
	
	
	
	
	
}
