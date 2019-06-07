package cn.wgh.excle;

import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POIUtils {
	/**
	 * 得到Excel，并解析内容
	 * 
	 * 对97-2003版本 使用HSSF解析 
	 * 对2007及以上版本 使用XSSF解析
	 * 
	 * 默认读取列数为第一行的列数，列数范围内的每行数据
	 * @param fileUrl
	 * @param removeEmptyRow
	 * @throws IOException
	 * @throws InvalidFormatException 
	 * @throws EncryptedDocumentException 
	 */
	@SuppressWarnings("deprecation")
	public static List<List<String>> getExcelAsFile(String fileUrl) throws IOException, EncryptedDocumentException, InvalidFormatException {
		// 1.得到Excel常用对象
		FileInputStream fileInputStream = new FileInputStream(fileUrl);
			Workbook wb = WorkbookFactory.create(fileInputStream);
			Sheet sheet = wb.getSheetAt(0);
			
			int pLength = sheet.getPhysicalNumberOfRows();
			System.out.println("物理行数 pLength:"+pLength);
			
			// 总行数
			int trLength = sheet.getLastRowNum();
			System.out.println("总行数："+trLength);
			trLength = trLength < pLength ? pLength : trLength + 1;
			// 4.得到Excel工作表的行
			Row row = sheet.getRow(0);
			// 总列数
			int tdLength = row.getLastCellNum();
			System.out.println("总列数："+tdLength);
			// 5.得到Excel工作表指定行的单元格
			Cell cell = row.getCell((short) 1);
			// 6.得到单元格样式
			CellStyle cellStyle = cell.getCellStyle();
			List<List<String>> trList = new ArrayList<List<String>>();
			
			System.out.println("-----------------------------------------------------------------------");
			for (int i = 0; i < trLength; i++) {
				// 得到Excel工作表的行
				Row row1 = sheet.getRow(i);
				List<String> tdList = new ArrayList<String>();
				for (int j = 0; j < tdLength; j++) {
					if(row1==null){
						break;
					}
					// 得到Excel工作表指定行的单元格
					Cell cell1 = row1.getCell(j);
					/**
					 * 为了处理：Excel异常Cannot get a text value from a numeric cell
					 * 将所有列中的内容都设置成String类型格式
					 */
					
					// 获得每一列中的值
					if (cell1 != null) {
//						if(HSSFDateUtil.isCellDateFormatted(cell1)){//如果是Excle的时间格式就转成文本格式
//							Date date = cell1.getDateCellValue();
//							SimpleDateFormat dateFormate = new SimpleDateFormat("yyyy-MM-dd");
//							String dateStr = dateFormate.format(date);
//							tdList.add(dateStr);
//							System.out.print("."+dateStr+ "\t\t");
//						}else{
							cell1.setCellType(Cell.CELL_TYPE_STRING);
							tdList.add(cell1.getStringCellValue());
							System.out.print(","+cell1.getStringCellValue()+ "\t\t");
//						}
						
					}else{
						tdList.add("");
						System.out.print("\t\t");
					}
//					System.out.print(cell1 + "\t\t");
//					System.out.print(cell1.getStringCellValue() + "\t\t");
				}
				trList.add(tdList);
				System.out.println();
			}
			return trList;
		}

	
	
	
	
	/**
	 * 得到Excel，并解析内容
	 * 
	 * 对97-2003版本 使用HSSF解析 
	 * 对2007及以上版本 使用XSSF解析
	 * 
	 * 默认读取列数为第一行的列数，列数范围内的每行数据，通过removeEmptyRow控制是否保留整行数据为空
	 * 
	 * @param fileUrl
	 * @param removeEmptyRow
	 * @throws IOException
	 * @throws InvalidFormatException 
	 * @throws EncryptedDocumentException 
	 */
	@SuppressWarnings("deprecation")
	public static List<List<String>> getExcelAsFile(String fileUrl, boolean removeEmptyRow) throws IOException, EncryptedDocumentException, InvalidFormatException {
		// 1.得到Excel常用对象
		FileInputStream fileInputStream = new FileInputStream(fileUrl);
			Workbook wb = WorkbookFactory.create(fileInputStream);
			Sheet sheet = wb.getSheetAt(0);
			
			int pLength = sheet.getPhysicalNumberOfRows();
			System.out.println("物理行数 pLength:"+pLength);
			
			// 总行数
			int trLength = sheet.getLastRowNum();
			System.out.println("总行数："+trLength);
			trLength = trLength < pLength ? pLength : trLength + 1;
			// 4.得到Excel工作表的行
			Row row = sheet.getRow(0);
			// 总列数
			int tdLength = row.getLastCellNum();
			System.out.println("总列数："+tdLength);
			// 5.得到Excel工作表指定行的单元格
			Cell cell = row.getCell((short) 1);
			// 6.得到单元格样式
			CellStyle cellStyle = cell.getCellStyle();
			List<List<String>> trList = new ArrayList<List<String>>();
			
			System.out.println("-----------------------------------------------------------------------");
			for (int i = 0; i < trLength; i++) {
				// 得到Excel工作表的行
				Row row1 = sheet.getRow(i);
				List<String> tdList = new ArrayList<String>();
				for (int j = 0; j < tdLength; j++) {
					if(row1==null){
						break;
					}
					// 得到Excel工作表指定行的单元格
					Cell cell1 = row1.getCell(j);
					/**
					 * 为了处理：Excel异常Cannot get a text value from a numeric cell
					 * 将所有列中的内容都设置成String类型格式
					 */
					
					// 获得每一列中的值
					if (cell1 != null) {
//						if(HSSFDateUtil.isCellDateFormatted(cell1)){//如果是Excle的时间格式就转成文本格式
//							Date date = cell1.getDateCellValue();
//							SimpleDateFormat dateFormate = new SimpleDateFormat("yyyy-MM-dd");
//							String dateStr = dateFormate.format(date);
//							tdList.add(dateStr);
//							System.out.print("."+dateStr+ "\t\t");
//						}else{
							cell1.setCellType(Cell.CELL_TYPE_STRING);
							tdList.add(cell1.getStringCellValue());
							System.out.print(","+cell1.getStringCellValue()+ "\t\t");
//						}
						
					}else{
						tdList.add("");
						System.out.print("\t\t");
					}
//					System.out.print(cell1 + "\t\t");
//					System.out.print(cell1.getStringCellValue() + "\t\t");
				}
				if(removeEmptyRow){
					if(tdList!=null && !tdList.isEmpty()){
						boolean addFlag = false;
						for (String str : tdList) {
							if(StringUtils.isNotEmpty(str)){
								addFlag = true;
								break;
							}
						}
						if(addFlag){
							trList.add(tdList);
						}
					}
				}else{
					trList.add(tdList);
				}
				System.out.println();
			}
			return trList;
		}

	
	
	
	
	
	/**
	 * 得到Excel，并解析内容
	 * 
	 * 
	 * 默认读取列数为第一行的列数，列数范围内的每行数据，通过removeEmptyRow控制是否保留整行数据为空 ,通过fillMaxtdLength控制是否将每行数据读取至最大列数
	 * 
	 * 对97-2003版本 使用HSSF解析 
	 * 对2007及以上版本 使用XSSF解析
	 * 
	 * @param fileUrl
	 * @param removeEmptyRow
	 * @throws IOException
	 * @throws InvalidFormatException 
	 * @throws EncryptedDocumentException 
	 */
	@SuppressWarnings("deprecation")
	public static List<List<String>> getExcelAsFile补全(String fileUrl, boolean removeEmptyRow, boolean fillMaxtdLength) throws IOException, EncryptedDocumentException, InvalidFormatException {
		// 1.得到Excel常用对象
		FileInputStream fileInputStream = new FileInputStream(fileUrl);
			Workbook wb = WorkbookFactory.create(fileInputStream);
			Sheet sheet = wb.getSheetAt(0);
			
			int pLength = sheet.getPhysicalNumberOfRows();
			System.out.println("物理行数 pLength:"+pLength);
			
			// 总行数
			int trLength = sheet.getLastRowNum();
			System.out.println("总行数："+trLength);
			trLength = trLength < pLength ? pLength : trLength + 1;
			// 4.得到Excel工作表的行
			Row row = sheet.getRow(0);
			// 总列数
			int tdLength = row.getLastCellNum();
//			System.out.println("总列数："+tdLength);

			int maxtdLength = 0;
			if (fillMaxtdLength) {
				for (int i = 0; i < trLength; i++) {
					if (i == 0 || (maxtdLength < sheet.getRow(i).getLastCellNum())) {
						maxtdLength = sheet.getRow(i).getLastCellNum();
					}
				}
				tdLength = maxtdLength;
			}
			
			// 5.得到Excel工作表指定行的单元格
			Cell cell = row.getCell((short) 1);
			// 6.得到单元格样式
			CellStyle cellStyle = cell.getCellStyle();
			List<List<String>> trList = new ArrayList<List<String>>();
			
			System.out.println("-----------------------------------------------------------------------");
			for (int i = 0; i < trLength; i++) {
				// 得到Excel工作表的行
				Row row1 = sheet.getRow(i);
				List<String> tdList = new ArrayList<String>();
				for (int j = 0; j < tdLength; j++) {
					if(j<row1.getLastCellNum()){
						if(row1==null){
							break;
						}
						// 得到Excel工作表指定行的单元格
						Cell cell1 = row1.getCell(j);
						/**
						 * 为了处理：Excel异常Cannot get a text value from a numeric cell
						 * 将所有列中的内容都设置成String类型格式
						 */
						
						// 获得每一列中的值
						if (cell1 != null) {
//							if(HSSFDateUtil.isCellDateFormatted(cell1)){//如果是Excle的时间格式就转成文本格式
//								Date date = cell1.getDateCellValue();
//								SimpleDateFormat dateFormate = new SimpleDateFormat("yyyy-MM-dd");
//								String dateStr = dateFormate.format(date);
//								tdList.add(dateStr);
//								System.out.print("."+dateStr+ "\t\t");
//							}else{
								CellStyle cellStyle2 = cell1.getCellStyle();
								cell1.setCellType(Cell.CELL_TYPE_STRING);
								
								tdList.add(cell1.getStringCellValue());
								System.out.print(","+cell1.getStringCellValue()+ "\t\t");
								
								
								
								
//								if(){
//									
//								}								
								
//								System.out.print(","+cell1.getStringCellValue()+ "cellStyle2:"+cellStyle2+"\t\t");
//							}
							
						}else{
							tdList.add("");
							System.out.print("\t\t");
						}
					}else if(fillMaxtdLength){
						tdList.add("");
					}
				}
				if(removeEmptyRow){
					if(tdList!=null && !tdList.isEmpty()){
						boolean addFlag = false;
						for (String str : tdList) {
							if(StringUtils.isNotEmpty(str)){
								addFlag = true;
								break;
							}
						}
						if(addFlag){
							trList.add(tdList);
						}
					}
				}else{
					trList.add(tdList);
				}
				System.out.println();
			}
			return trList;
		}
	
	
	
	
	@SuppressWarnings("deprecation")
	public static List<List<String>> getExcelAsFile补全2(String fileUrl, boolean removeEmptyRow, boolean fillMaxtdLength) throws IOException, EncryptedDocumentException, InvalidFormatException {
		// 1.得到Excel常用对象
		FileInputStream fileInputStream = new FileInputStream(fileUrl);
//			Workbook wb = WorkbookFactory.create(fileInputStream);
			XSSFWorkbook wb =  new XSSFWorkbook(fileInputStream); 
		
			
			//==========
			//创建workBook 
//			 HSSFWorkbook wb222 = new HSSFWorkbook(fileInputStream); 
			  //创建一个样式 
//			  HSSFCellStyle cellStyle222 =  wb222.createCellStyle(); 
//			 //创建一个DataFormat对象 
//			 HSSFDataFormat format =  wb222.createDataFormat(); 
//			 //这样才能真正的控制单元格格式，@就是指文本型，具体格式的定义还是参考上面的原文吧 
//			 cellStyle222.setDataFormat(format.getFormat("@")); 
			 
			 //具体如何创建cell就省略了，最后设置单元格的格式这样写 
//			 cell.setCellStyle(cellStyle222);
			//==========
			//==========
			//创建workBook 
//			 HSSFWorkbook wb222 = new HSSFWorkbook(fileInputStream); 
			//创建一个样式 
			XSSFCellStyle cellStyle222 =  wb.createCellStyle(); 
//			 //创建一个DataFormat对象 
//			XSSFDataFormat format =  wb.createDataFormat(); 
//			 //这样才能真正的控制单元格格式，@就是指文本型，具体格式的定义还是参考上面的原文吧 
//			 cellStyle222.setDataFormat(format.getFormat("¥#,##0")); 
			
			//具体如何创建cell就省略了，最后设置单元格的格式这样写 
//			 cell.setCellStyle(cellStyle222);
			//==========
			
			
			 XSSFSheet sheet =  wb.getSheetAt(0);
			
			int pLength = sheet.getPhysicalNumberOfRows();
			System.out.println("物理行数 pLength:"+pLength);
			
			// 总行数
			int trLength = sheet.getLastRowNum();
			System.out.println("总行数："+trLength);
			trLength = trLength < pLength ? pLength : trLength + 1;
			// 4.得到Excel工作表的行
			Row row = sheet.getRow(0);
			// 总列数
			int tdLength = row.getLastCellNum();
//			System.out.println("总列数："+tdLength);

			int maxtdLength = 0;
			if (fillMaxtdLength) {
				for (int i = 0; i < trLength; i++) {
					if (i == 0 || (maxtdLength < sheet.getRow(i).getLastCellNum())) {
						maxtdLength = sheet.getRow(i).getLastCellNum();
					}
				}
				tdLength = maxtdLength;
			}
			
			// 5.得到Excel工作表指定行的单元格
			Cell cell = row.getCell((short) 1);
			// 6.得到单元格样式
			CellStyle cellStyle = cell.getCellStyle();
			List<List<String>> trList = new ArrayList<List<String>>();
			
			System.out.println("-----------------------------------------------------------------------");
			for (int i = 0; i < trLength; i++) {
				// 得到Excel工作表的行
				XSSFRow row1 =  sheet.getRow(i);
				List<String> tdList = new ArrayList<String>();
				for (int j = 0; j < tdLength; j++) {
					if (j < row1.getLastCellNum()) {
						if (row1 == null) {
							break;
						}
						// 得到Excel工作表指定行的单元格
						XSSFCell cell1 = row1.getCell(j);
						/**
						 * 为了处理：Excel异常Cannot get a text value from a numeric cell
						 * 将所有列中的内容都设置成String类型格式
						 */
						// 获得每一列中的值
						if (cell1 != null) {
							short  celldataFormat = cell1.getCellStyle().getDataFormat();
							if (celldataFormat==14||celldataFormat == 31 || celldataFormat == 57 || celldataFormat == 58 ) {// 如果是Excle的时间格式就转成文本格式
								tdList.add(changeDateCellToString(cell1));
								System.out.print("," + changeDateCellToString(cell1) + "\t\t");
							}else if(celldataFormat == 20 || celldataFormat == 32){
								tdList.add(changeTimeCellToString(cell1));
								System.out.print("," + changeTimeCellToString(cell1) + "\t\t");
							}else if(celldataFormat == 22){
								tdList.add(changeDateTimeCellToString(cell1));
								System.out.print("," + changeDateTimeCellToString(cell1) + "\t\t");
							}else if (cell1.getCellType() == XSSFCell.CELL_TYPE_STRING) {
								tdList.add(cell1.getStringCellValue());
								System.out.print("," + cell1.getStringCellValue() + "\t\t");
							} else if (cell1.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
								tdList.add(changeNUMCellToString(cell1));
								System.out.print("," + changeNUMCellToString(cell1) + "\t\t");
							}else{
								cell1.setCellType(Cell.CELL_TYPE_STRING);
								tdList.add(cell1.getStringCellValue());
								System.out.print("," + cell1.getStringCellValue() + "\t\t");
							}
						} else {
							tdList.add("");
							System.out.print("\t\t");
						}
					} else if (fillMaxtdLength) {
						tdList.add("");
					}
				}
				if(removeEmptyRow){
					if(tdList!=null && !tdList.isEmpty()){
						boolean addFlag = false;
						for (String str : tdList) {
							if(StringUtils.isNotEmpty(str)){
								addFlag = true;
								break;
							}
						}
						if(addFlag){
							trList.add(tdList);
						}
					}
				}else{
					trList.add(tdList);
				}
				System.out.println();
			}
			return trList;
		}
	
	
	@SuppressWarnings("deprecation")
	public static List<List<String>> getExcelAsFile补全3(String fileUrl, boolean removeEmptyRow, boolean fillMaxtdLength) throws IOException, EncryptedDocumentException, InvalidFormatException {
		// 1.得到Excel常用对象
		FileInputStream fileInputStream = new FileInputStream(fileUrl);
			Workbook wb = WorkbookFactory.create(fileInputStream);
			Sheet sheet = wb.getSheetAt(0);
			
			int pLength = sheet.getPhysicalNumberOfRows();
			System.out.println("物理行数 pLength:"+pLength);
			
			// 总行数
			int trLength = sheet.getLastRowNum();
			System.out.println("总行数："+trLength);
			trLength = trLength < pLength ? pLength : trLength + 1;
			// 4.得到Excel工作表的行
			Row row = sheet.getRow(0);
			// 总列数
			int tdLength = row.getLastCellNum();
//			System.out.println("总列数："+tdLength);

			int maxtdLength = 0;
			if (fillMaxtdLength) {
				for (int i = 0; i < trLength; i++) {
					if (i == 0 || (maxtdLength < sheet.getRow(i).getLastCellNum())) {
						maxtdLength = sheet.getRow(i).getLastCellNum();
					}
				}
				tdLength = maxtdLength;
			}
			
			// 5.得到Excel工作表指定行的单元格
			Cell cell = row.getCell((short) 1);
			// 6.得到单元格样式
			CellStyle cellStyle = cell.getCellStyle();
			List<List<String>> trList = new ArrayList<List<String>>();
			
			System.out.println("-----------------------------------------------------------------------");
			for (int i = 0; i < trLength; i++) {
				// 得到Excel工作表的行
				Row row1 = sheet.getRow(i);
				List<String> tdList = new ArrayList<String>();
				for (int j = 0; j < tdLength; j++) {
					if(j<row1.getLastCellNum()){
						if(row1==null){
							break;
						}
						// 得到Excel工作表指定行的单元格
						Cell cell1 = row1.getCell(j);
						/**
						 * 为了处理：Excel异常Cannot get a text value from a numeric cell
						 * 将所有列中的内容都设置成String类型格式
						 */
						
						// 获得每一列中的值
						if (cell1 != null) {
//							if(HSSFDateUtil.isCellDateFormatted(cell1)){//如果是Excle的时间格式就转成文本格式
//								Date date = cell1.getDateCellValue();
//								SimpleDateFormat dateFormate = new SimpleDateFormat("yyyy-MM-dd");
//								String dateStr = dateFormate.format(date);
//								tdList.add(dateStr);
//								System.out.print("."+dateStr+ "\t\t");
//							}else{
								CellStyle cellStyle2 = cell1.getCellStyle();
								cell1.setCellType(Cell.CELL_TYPE_STRING);
								

								if(cell1.getCellType()==Cell.CELL_TYPE_STRING){
									tdList.add(cell1.getStringCellValue() +"]]]]"+cell1.getCellType());
									System.out.print(","+cell1.getStringCellValue()+ "\t\t");
								}
								if(cell1.getCellType()==Cell.CELL_TYPE_NUMERIC){
									tdList.add(String.valueOf(cell1.getNumericCellValue()+"]]]]"+cell1.getCellType()));
									System.out.print(","+String.valueOf(cell1.getNumericCellValue())+ "\t\t");
//									System.out.print(","+cell1.getNumericCellValue()+ "\t\t");
								}
								
//							}
							
						}else{
							tdList.add("");
							System.out.print("\t\t");
						}
					}else if(fillMaxtdLength){
						tdList.add("");
					}
				}
				if(removeEmptyRow){
					if(tdList!=null && !tdList.isEmpty()){
						boolean addFlag = false;
						for (String str : tdList) {
							if(StringUtils.isNotEmpty(str)){
								addFlag = true;
								break;
							}
						}
						if(addFlag){
							trList.add(tdList);
						}
					}
				}else{
					trList.add(tdList);
				}
				System.out.println();
			}
			return trList;
		}

	
//	StringUtils.strip();
	
	public static String changeDateCellToString(XSSFCell cell) {
		// 如果是Excle的日期格式就转成文本格式
		Date date = cell.getDateCellValue();
		SimpleDateFormat dateFormate = new SimpleDateFormat("yyyy-MM-dd");
		return dateFormate.format(date);
	}
	
	
	public static String changeTimeCellToString(XSSFCell cell) {
		// 如果是Excle的时间格式就转成文本格式
		Date date = cell.getDateCellValue();
		SimpleDateFormat dateFormate = new SimpleDateFormat("HH:mm:ss");
		return dateFormate.format(date);
	}
	
	
	public static String changeDateTimeCellToString(XSSFCell cell) {
		// 如果是Excle的日期时间格式就转成文本格式
		Date date = cell.getDateCellValue();
		SimpleDateFormat dateFormate = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		return dateFormate.format(date);
	}
	
	
	public static String changeNUMCellToString(XSSFCell cell) {
		String res = "";
		
		double numCellVal = cell.getNumericCellValue();
		
		BigDecimal bdmal = new BigDecimal(numCellVal);
		
		if(overDecimalPoint(numCellVal,2) || !isPointNum(cell)){
			DecimalFormat df =new DecimalFormat();
			res = df.format(bdmal);
		}else{
			DecimalFormat df =new DecimalFormat(",###.00");
			res = df.format(bdmal);
		}
		return res;
	}
	
	
	private static boolean overDecimalPoint(double d, int i) {
		if (i < (StringUtils.substringAfterLast(String.valueOf(d), ".")).length()) {
			return true;
		}
		return false;
	}
	
	private static boolean isPointNum(XSSFCell cell) {
		if(cell.getCellType() == XSSFCell.CELL_TYPE_FORMULA){
			HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
			String cellFormatted = dataFormatter.formatCellValue(cell);
			System.err.println();
			System.err.println("============="+cellFormatted);
		}else{
			HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
			String cellFormatted = dataFormatter.formatCellValue(cell);
			System.err.println();
			System.err.println("============="+cellFormatted);
			
			return StringUtils.contains(cellFormatted, ".");
			
		}
		return true;
		
		
	}
	
	
	public static String changeNUMCellToString2(XSSFCell cell ,XSSFCellStyle cellStyle) {
		String cellNumVal = String.valueOf(cell.getNumericCellValue());
		XSSFCell cellStr = cell;
		cellStr.setCellType(Cell.CELL_TYPE_STRING);
		String cellStrVal = cellStr.getStringCellValue();
		
		cell.setCellType(Cell.CELL_TYPE_NUMERIC);
//		                    String.valueOf(cell1.getNumericCellValue()
		if (cellNumVal.contains(".0") && cellStrVal.equals(cellNumVal.replace(".0", ""))) {
			return cellStrVal;
		}
		
		
		
//        cellStyle.setDataFormat(XSSFDataFormat.getFormat("0.00"));
        cell.setCellStyle(cellStyle);
		
		
		return cellNumVal;
	}
	
	
	
	
	
	public static String changeCellToString(XSSFCell cell) {
		String returnValue = "";
		if (null != cell) {
			switch (cell.getCellType()) {
			case XSSFCell.CELL_TYPE_NUMERIC: // 数字
				Double doubleValue = cell.getNumericCellValue();
				String str = doubleValue.toString();
				if (str.contains(".0")) {
					str = str.replace(".0", "");
				}
//				Integer intValue = Integer.parseInt(str);
//				returnValue = intValue.toString();
				returnValue = str;
				break;
			case XSSFCell.CELL_TYPE_STRING: // 字符串
				returnValue = cell.getStringCellValue();
				break;
			case XSSFCell.CELL_TYPE_BOOLEAN: // 布尔
				Boolean booleanValue = cell.getBooleanCellValue();
				returnValue = booleanValue.toString();
				break;
			case XSSFCell.CELL_TYPE_BLANK: // 空值
				returnValue = "";
				break;

			case XSSFCell.CELL_TYPE_FORMULA: // 公式
				returnValue = cell.getCellFormula();
				break;
			case XSSFCell.CELL_TYPE_ERROR: // 故障
				returnValue = "";
				break;
			default:
				System.out.println("未知类型");
				break;
			}
		}
		return returnValue;
	}
	
	
	
	@SuppressWarnings("deprecation")
	public static List<List<String>> getExcelAsFile补全20170823(String fileUrl, boolean removeEmptyRow, boolean fillMaxtdLength) throws IOException, EncryptedDocumentException, InvalidFormatException {
		// 1.得到Excel常用对象
		FileInputStream fileInputStream = new FileInputStream(fileUrl);
//			Workbook wb = WorkbookFactory.create(fileInputStream);
			XSSFWorkbook wb =  new XSSFWorkbook(fileInputStream); 
		
			
			//==========
			//创建workBook 
//			 HSSFWorkbook wb222 = new HSSFWorkbook(fileInputStream); 
			  //创建一个样式 
//			  HSSFCellStyle cellStyle222 =  wb222.createCellStyle(); 
//			 //创建一个DataFormat对象 
//			 HSSFDataFormat format =  wb222.createDataFormat(); 
//			 //这样才能真正的控制单元格格式，@就是指文本型，具体格式的定义还是参考上面的原文吧 
//			 cellStyle222.setDataFormat(format.getFormat("@")); 
			 
			 //具体如何创建cell就省略了，最后设置单元格的格式这样写 
//			 cell.setCellStyle(cellStyle222);
			//==========
			//==========
			//创建workBook 
//			 HSSFWorkbook wb222 = new HSSFWorkbook(fileInputStream); 
			//创建一个样式 
			XSSFCellStyle cellStyle222 =  wb.createCellStyle(); 
//			 //创建一个DataFormat对象 
//			XSSFDataFormat format =  wb.createDataFormat(); 
//			 //这样才能真正的控制单元格格式，@就是指文本型，具体格式的定义还是参考上面的原文吧 
//			 cellStyle222.setDataFormat(format.getFormat("¥#,##0")); 
			
			//具体如何创建cell就省略了，最后设置单元格的格式这样写 
//			 cell.setCellStyle(cellStyle222);
			//==========
			
			
			 XSSFSheet sheet =  wb.getSheetAt(0);
			
			int pLength = sheet.getPhysicalNumberOfRows();
			System.out.println("物理行数 pLength:"+pLength);
			
			// 总行数
			int trLength = sheet.getLastRowNum();
			System.out.println("总行数："+trLength);
			trLength = trLength < pLength ? pLength : trLength + 1;
			// 4.得到Excel工作表的行
			Row row = sheet.getRow(0);
			// 总列数
			int tdLength = row.getLastCellNum();
//			System.out.println("总列数："+tdLength);

			int maxtdLength = 0;
			if (fillMaxtdLength) {
				for (int i = 0; i < trLength; i++) {
					if (i == 0 || (maxtdLength < sheet.getRow(i).getLastCellNum())) {
						maxtdLength = sheet.getRow(i).getLastCellNum();
					}
				}
				tdLength = maxtdLength;
			}
			
			// 5.得到Excel工作表指定行的单元格
			Cell cell = row.getCell((short) 1);
			// 6.得到单元格样式
			CellStyle cellStyle = cell.getCellStyle();
			List<List<String>> trList = new ArrayList<List<String>>();
			
			System.out.println("-----------------------------------------------------------------------");
			for (int i = 0; i < trLength; i++) {
				// 得到Excel工作表的行
				XSSFRow row1 =  sheet.getRow(i);
				List<String> tdList = new ArrayList<String>();
				for (int j = 0; j < tdLength; j++) {
					if (j < row1.getLastCellNum()) {
						if (row1 == null) {
							break;
						}
						// 得到Excel工作表指定行的单元格
						XSSFCell cell1 = row1.getCell(j);
						/**
						 * 为了处理：Excel异常Cannot get a text value from a numeric cell
						 * 将所有列中的内容都设置成String类型格式
						 */
						// 获得每一列中的值
						if (cell1 != null) {
							short  celldataFormat = cell1.getCellStyle().getDataFormat();
							if (celldataFormat==14||celldataFormat == 31 || celldataFormat == 57 || celldataFormat == 58 ) {// 如果是Excle的时间格式就转成文本格式
								tdList.add(changeDateCellToString(cell1));
								System.out.print("," + changeDateCellToString(cell1) + "\t\t");
							}else if(celldataFormat == 20 || celldataFormat == 32){
								tdList.add(changeTimeCellToString(cell1));
								System.out.print("," + changeTimeCellToString(cell1) + "\t\t");
							}else if(celldataFormat == 22){
								tdList.add(changeDateTimeCellToString(cell1));
								System.out.print("," + changeDateTimeCellToString(cell1) + "\t\t");
							}else if (cell1.getCellType() == XSSFCell.CELL_TYPE_STRING) {
								tdList.add(cell1.getStringCellValue());
								System.out.print("," + cell1.getStringCellValue() + "\t\t");
							}else if (cell1.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
								tdList.add(changeNUMCellToString(cell1));
//								System.out.print("," + changeNUMCellToString(cell1) + "\t\t");
							}else if(cell1.getCellType() == XSSFCell.CELL_TYPE_FORMULA){
								try {  
									tdList.add(cell1.getStringCellValue());  
									System.out.print("," + cell1.getStringCellValue() + "\t\t");
								} catch (IllegalStateException e) {  
									tdList.add(changeNUMCellToString(cell1));
//									System.out.print("," + changeNUMCellToString(cell1) + "\t\t");
								}  
							}else{
								cell1.setCellType(Cell.CELL_TYPE_STRING);
								tdList.add(cell1.getStringCellValue());
								System.out.print("," + cell1.getStringCellValue() + "\t\t");
							}
						} else {
							tdList.add("");
							System.out.print("\t\t");
						}
					} else if (fillMaxtdLength) {
						tdList.add("");
					}
				}
				if(removeEmptyRow){
					if(tdList!=null && !tdList.isEmpty()){
						boolean addFlag = false;
						for (String str : tdList) {
							if(StringUtils.isNotEmpty(str)){
								addFlag = true;
								break;
							}
						}
						if(addFlag){
							trList.add(tdList);
						}
					}
				}else{
					trList.add(tdList);
				}
				System.out.println();
			}
			return trList;
		}
}
