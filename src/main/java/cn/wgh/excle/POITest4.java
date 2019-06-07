package cn.wgh.excle;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;
import org.junit.Test;

public class POITest4 {
	@Test
	public void 删除已存在的Excle的指定列内容_写入新类容() throws Exception {
		FileInputStream fs = new FileInputStream("E:/test/xlsTest/POITest/干部资料.xls");// 获取
		POIFSFileSystem ps = new POIFSFileSystem(fs); // 使用POI提供的方法得到excel的信息
		HSSFWorkbook wb = new HSSFWorkbook(ps);
		HSSFSheet sheet = wb.getSheetAt(0); // 获取到工作表，因为一个excel可能有多个工作表
//		HSSFSheet sheet = wb.getSheet("Sheet2"); // 根据名字获取到工作表
		
		int lastRowNum = sheet.getLastRowNum();
		System.out.println("lastRowNum:"+lastRowNum);
		//指定要删除的列
		List<Integer> cellnumList = new ArrayList<Integer>();
		cellnumList.add(0);
		cellnumList.add(1);
		for (Integer cellnum : cellnumList) {
			System.err.println("===============================================================");
			for (int i = 0; i < lastRowNum; i++) {
				System.err.println("================"+i);
				HSSFRow row = sheet.getRow(i);
				HSSFCell cell = row.getCell(cellnum);
				
				int rowNum = row.getRowNum();
				HSSFRow cellRow = cell.getRow();
				int rowIndex = cell.getRowIndex();
				System.err.println("rowNum:"+rowNum);
				System.err.println("rowIndex:"+rowIndex);
				String stringCellValue = cell.getStringCellValue();
				System.err.println("stringCellValue:"+stringCellValue);
				
				if("单位名称".equals(cell.getStringCellValue())){
//					row.removeCell(cell);
//					cell.setCellValue("哈哈哈");
					row.createCell(cellnum).setCellValue("哈哈哈");
					CellStyle cellStyle = cell.getCellStyle();
					row = sheet.createRow((short) (sheet.getLastRowNum() + 1)); // 在现有行号后追加数据
					row.createCell(0).setCellValue("旺财1"); // 设置第一个（从0开始）单元格的数据
					row.createCell(1).setCellValue(241); // 设置第二个（从0开始）单元格的数据
					row.createCell(2).setCellValue("啦啦啦1111！");
					row = sheet.createRow((short) (rowIndex + 1)); // 在现有行号后追加数据
					row.createCell(0).setCellValue("旺财"); // 设置第一个（从0开始）单元格的数据
					row.createCell(1).setCellValue(24); // 设置第二个（从0开始）单元格的数据
					row.createCell(2).setCellValue("啦啦啦！");
					
//					row.
//					sheet.
				}
				
				
				
				if("${name}".equals(cell.getStringCellValue())){
					row.removeCell(cell);
//					cell.setCellValue("哈哈哈");
					row.createCell(cellnum).setCellValue("哈哈哈");
				}
				
				if("abc".equals(cell.getStringCellValue())){
					row.removeCell(cell);
					row.createCell(cellnum).setCellValue("哈哈哈123");
				}
//				if (cell != null) {//判断cell为空值:if(cell==null||cell.equals("")||cell.getCellType() ==HSSFCell.CELL_TYPE_BLANK)
//					row.removeCell(cell);
//				}
			}
		}
		 
//		HSSFCellStyle cellStyle = wb.createCellStyle();    
//		 一、设置背景色:  
//		  
//		  
//		cellStyle.setFillForegroundColor((short) 13);// 设置背景色    
//		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);    
//		二、设置边框:  
//		  
//		  
//		cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框    
//		cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框    
//		cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框    
//		cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框    
//		三、设置居中:  
		  
		  
//		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中   
		
//		List<TreePojo> treePojoDatas =getTreePojoDatas();
////		
//		int size = treePojoDatas.size();
//		if (size > 0) {
//			int datanum = 0;
//			for (int rownum = 0; rownum <= lastRowNum; rownum++) {
//				HSSFRow row = sheet.getRow(rownum);
//				for (Integer cellnum : cellnumList) {
//					if (datanum < size) {
//						row.createCell(cellnum).setCellValue(treePojoDatas.get(datanum++).getName()); // 设置第一个（从0开始）单元格的数据
//					}
//				}
//			}
//		}
		FileOutputStream out = new FileOutputStream("E:/test/xlsTest/POITest/干部资料修改"+System.currentTimeMillis()+".xls");//// 向xls中写数据

		out.flush();
		wb.write(out);
		out.close();
	}
	
	
	
	public List<TreePojo> getTreePojoDatas(){
		List<TreePojo> list = new ArrayList<TreePojo>();
		list.add(new TreePojo("101", "0", "四川省", "scs"));
		list.add(new TreePojo("102", "0", "湖北省", "hbs"));
		list.add(new TreePojo("201", "101", "成都市", "cds"));
		list.add(new TreePojo("204", "201", "武汉市", "whs"));
		list.add(new TreePojo("203", "101", "绵阳市", "mys"));
		list.add(new TreePojo("205", "201", "襄阳市", "xys"));
		list.add(new TreePojo("202", "101", "乐山市", "lss"));
		list.add(new TreePojo("206", "201", "十堰市", "sys"));
		return list;
	}
	
}
