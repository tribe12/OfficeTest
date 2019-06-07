package cn.wgh.excle;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class POITest {
	@SuppressWarnings("unused")  
	@Test
	public void createExcel() throws Exception {
		 //生成Workbook  
		HSSFWorkbook wb = new HSSFWorkbook();  
		  
//		//添加Worksheet（不添加sheet时生成的xls文件打开时会报错）  
		Sheet sheet1 = wb.createSheet();  
//		Sheet sheet2 = wb.createSheet();  
//		Sheet sheet3 = wb.createSheet("new sheet");  
//		Sheet sheet4 = wb.createSheet("rensanning");  
  
		//保存为Excel文件  
		FileOutputStream out = null;  
		  
		try {  
		    out = new FileOutputStream("E:\\test\\xlsTest\\我是POI生成的100000.xls");  
		    wb.write(out);        
		} catch (IOException e) {  
		    System.out.println(e.toString());  
		} finally {  
		    try {  
		        out.close();  
		    } catch (IOException e) {  
		        System.out.println(e.toString());  
		    }  
		}
	}
	
	
	@SuppressWarnings("unused")  
	@Test
	public void createExcel2() throws Exception {
		//生成Workbook  
		XSSFWorkbook wb = new XSSFWorkbook();    
		
//		//添加Worksheet（不添加sheet时生成的xls文件打开时会报错）  
		Sheet sheet1 = wb.createSheet();  
		Sheet sheet2 = wb.createSheet();  
		Sheet sheet3 = wb.createSheet("new sheet");  
		Sheet sheet4 = wb.createSheet("rensanning");  
		
		//保存为Excel文件  
		FileOutputStream out = null;  
		
		try {  
			out = new FileOutputStream("E:\\test\\xlsTest\\我是POI生成的1.xls");  
			wb.write(out);        
		} catch (IOException e) {  
			System.out.println(e.toString());  
		} finally {  
			try {  
				out.close();  
			} catch (IOException e) {  
				System.out.println(e.toString());  
			}  
		}
	}
	
	
	/**
	 * 得到Excel，并解析内容 
	 * 
	 * 对97-2003版本 使用HSSF解析
	 * 对2007及以上版本 使用XSSF解析
	 * @throws Exception
	 */
	@Test
	public void getExcelAsFile() throws Exception {
		String fileUrl = "E:/test/xlsTest/XSSF/111.xlsx";
//		fileUrl = "E:/test/xlsTest/XSSF/stuTest2.xls";
//		fileUrl = "E:/test/xlsTest/XSSF/stuTest3.xls";
		fileUrl = "C:/Users/Lenovo/Desktop/Base/人员信息-王国辉 - 副本.xls";
		List<List<String>> excelAsFile = POIUtils.getExcelAsFile(fileUrl);
		excelAsFile.remove(0);
		System.err.println("=====================================================");
		System.err.println(excelAsFile);
		System.err.println("=====================================================");
		for (List<String> list : excelAsFile) {
			System.err.println(list);
		}
		System.err.println("=====================================================");
		
		for (List<String> list : excelAsFile) {
			for (String string : list) {
				System.err.print(string + "\t\t");
			}
			System.err.println();
		}
	}
	
	
	
	
	@Test
	public void getExcelAsFileTest() throws Exception {
//		String fileUrl = "C:\\Users\\Lenovo\\Desktop\\天津团委的工资表test.xlsx";
//		String fileUrl = "C:\\Users\\Lenovo\\Desktop\\天津团委的工资表test - 副本.xlsx";
//		String fileUrl = "C:\\Users\\Lenovo\\Desktop\\天津团委的工资表test - 副本 - 副本.xlsx";
//		String fileUrl = "C:\\Users\\Lenovo\\Desktop\\工资表 (补发201703奖金)1000多条数据.xlsx";
		String fileUrl = "C:\\Users\\Lenovo\\Desktop\\ccccccccccccccccccccccccc\\工资表 (2222.8位数超长)10条.xls";
//		String fileUrl = "C:\\Users\\Lenovo\\Desktop\\ccccccccccccccccccccccccc\\工资表 (2222.8位数超长)10条WPS.xls";
//		String fileUrl = "C:\\Users\\Lenovo\\Desktop\\ccccccccccccccccccccccccc\\工资表 (2222.8位数超长)10条.xlsx";
//		String fileUrl = "C:\\Users\\Lenovo\\Desktop\\ccccccccccccccccccccccccc\\工资表 (2222.8位数超长)10条WPS.xlsx";
//		String fileUrl = "C:\\Users\\Lenovo\\Desktop\\ccccccccccccccccccccccccc\\1111111111青少年社工代发工资（8月工资表） (自动保存的).xlsx";
//		String fileUrl = "C:\\Users\\Lenovo\\Desktop\\ccccccccccccccccccccccccc\\111WPS1111111青少年社工代发工资（8月工资表） (自动保存的) - 副本.xlsx";
//		String fileUrl = "C:\\Users\\Lenovo\\Desktop\\E12.xlsx";
//		List<List<String>> excelAsFile = POIUtils.getExcelAsFile补全(fileUrl,true,true);
//		List<List<String>> excelAsFile = POIUtils.getExcelAsFile补全2(fileUrl,true,true);
//		List<List<String>> excelAsFile = POIUtils.getExcelAsFile补全3(fileUrl,true,true);
//		List<List<String>> excelAsFile = POIUtils.getExcelAsFile(fileUrl,false);
		List<List<String>> excelAsFile = POIUtils.getExcelAsFile补全20170823(fileUrl,true,true);
//		excelAsFile.remove(0);
		System.err.println("=========================");
		System.err.println(excelAsFile);
		System.err.println("=========================");
		for (List<String> list : excelAsFile) {
			System.err.println(list);
		}
		System.err.println();
		System.err.println();
		System.err.println();
		System.err.println();
		System.err.println("++++++++++++++++++++++++++++++++++++++++++++++++++++++++");
		
//		removeHeaders(excelAsFile, getRemoveHeaders(excelAsFile));
//		removeHeaders2(excelAsFile, getRemoveHeaders2(excelAsFile));
		
//		System.err.println("+++++++++++++++++++++++++");
//		System.err.println(excelAsFile);
//		System.err.println("++++++++++++遍历+++++++++++++");
//		for (List<String> list : excelAsFile) {
//			System.err.println(list);
//		}
	}
	
//	removeHeader
	
	public void removeHeaders(List<List<String>> excelAsFile, List<RemoveHeader> removeHeaders) {
		if (removeHeaders != null && !removeHeaders.isEmpty()) {
			List<List<String>> remList = new ArrayList<List<String>>();
			if (excelAsFile.size() == removeHeaders.size()) {
				int titleSort = 0;
				int minCount = 0;
				for (int i = 0; i < removeHeaders.size(); i++) {
					//------------
					List<String> empHeaderList = removeHeaders.get(i).getEmpHeaderList();
					int sort = removeHeaders.get(i).getSort();
					int emptyCount = removeHeaders.get(i).getEmptyCount();
					//------------
					if(i == 0 || (minCount > removeHeaders.get(i).getEmptyCount())){
						minCount = removeHeaders.get(i).getEmptyCount();
						titleSort = removeHeaders.get(i).getSort();
					}
				}
				for (int i = 0; i < titleSort; i++) {
					remList.add(removeHeaders.get(i).getEmpHeaderList());
				}
			} else {
				for (RemoveHeader remHeader : removeHeaders) {
					remList.add(remHeader.getEmpHeaderList());
				}
			}

			excelAsFile.removeAll(remList);
		}
	}
	
	public List<RemoveHeader> getRemoveHeaders(List<List<String>> rowList) {
		List<RemoveHeader> removeHeaders = new ArrayList<RemoveHeader>();
		if (rowList != null && !rowList.isEmpty()) {
			for (int sort = 0; sort < rowList.size(); sort++) {
				if(rowList.get(sort)!=null && !rowList.get(sort).isEmpty()){
					boolean addHeaderFlag = false;
					boolean allFullFlag = true;
					int emptyCount = 0;
					for (int i = 0; i < rowList.get(sort).size(); i++) {
						String string = rowList.get(sort).get(i);
						if(StringUtils.isEmpty(rowList.get(sort).get(i))){
							emptyCount++;
							addHeaderFlag = true;
							allFullFlag = false;
						}
					}
					if(allFullFlag){
						return removeHeaders;
					}
					if(addHeaderFlag ){
						removeHeaders.add(new RemoveHeader(sort,emptyCount,rowList.get(sort)));
					}
				}
			}
		}
		return removeHeaders;
	}
	
	
	
	public void removeHeaders2( List<List<String>> excelAsFile , List<List<String>>  removeHeaders) {
		if(removeHeaders!=null && !removeHeaders.isEmpty()){
				excelAsFile.removeAll(removeHeaders);
		}
		
		
	}
	public List<List<String>> getRemoveHeaders2(List<List<String>> rowList) {
		List<List<String>> removeHeaders = new ArrayList<List<String>>();
		if (rowList != null && !rowList.isEmpty()) {
			for (List<String> cellList : rowList) {
				if (cellList != null && !cellList.isEmpty()) {
					boolean addHeaderFlag = false;
					boolean allFullFlag = true;
					for (int i = 0; i < cellList.size(); i++) {
//						addHeaderFlag = false;
//						allFullFlag = true;
						if (StringUtils.isEmpty(cellList.get(i))) {
							addHeaderFlag = true;
							allFullFlag = false;
						}
					}
					if (allFullFlag) {
						return removeHeaders;
					}
					if (addHeaderFlag) {
						removeHeaders.add(cellList);
					}
				}
			}
		}
		return removeHeaders;
	}
	
	
	
	
	
	@Test
	public void testName3333333() throws Exception {
		List<Object> list = new ArrayList<Object>();
		for (int i = 0; i < 10; i++) {
			list.add(i);
		}
		System.out.println(list);
		ttt(list);
		System.out.println(list);
		
	}	
	
	
	public void  ttt(List<Object> list ){
		list.remove(0);
		list.remove(3);
	}
	
}
