package cn.wgh.excle;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.springframework.web.multipart.MultipartFile;

public class SalaryPOIUtils {

	/**
	 * 得到Excel，并解析内容
	 * 
	 * @param fileUrl
	 * @param removeEmptyRow
	 * @throws IOException
	 * @throws InvalidFormatException
	 * @throws EncryptedDocumentException
	 */
	public static List<List<String>> getExcelAsFile0823解决公式及小数点问题(
//			MultipartFile mulfile,
			String fileUrl, boolean removeEmptyRow,
			boolean fillMaxtdLength) throws IOException, EncryptedDocumentException, InvalidFormatException {
		// 1.得到Excel常用对象
		InputStream fileInputStream = null;
		if (StringUtils.isNotEmpty(fileUrl)) {
			fileInputStream = new FileInputStream(fileUrl);
		} else {
//			fileInputStream = mulfile.getInputStream();
		}

		XSSFWorkbook wb = new XSSFWorkbook(fileInputStream); 
		XSSFSheet sheet = wb.getSheetAt(0);

		int pLength = sheet.getPhysicalNumberOfRows();

		// 总行数
		int trLength = sheet.getLastRowNum();
		trLength = trLength < pLength ? pLength : trLength + 1;
		// 4.得到Excel工作表的行
		XSSFRow row = sheet.getRow(0);
		// 总列数
		int tdLength = row.getLastCellNum();

		int maxtdLength = 0;
		if (fillMaxtdLength) {
			for (int i = 0; i < trLength; i++) {
				if (i == 0 || (maxtdLength < sheet.getRow(i).getLastCellNum())) {
					maxtdLength = sheet.getRow(i).getLastCellNum();
				}
			}
			tdLength = maxtdLength;
		}

		List<List<String>> trList = new ArrayList<List<String>>();
		for (int i = 0; i < trLength; i++) {
			// 得到Excel工作表的行
			XSSFRow row1 = sheet.getRow(i);
			List<String> tdList = new ArrayList<String>();
			for (int j = 0; j < tdLength; j++) {
				if (j < row1.getLastCellNum()) {
					if (row1 == null) {
						break;
					}
					// 得到Excel工作表指定行的单元格
					XSSFCell cell1 = row1.getCell(j);
					if (cell1 != null) {
						// 获得每一列中的值。。。
						short  celldataFormat = cell1.getCellStyle().getDataFormat();
						if (celldataFormat == 14 || celldataFormat == 31 || celldataFormat == 57 || celldataFormat == 58) {// 如果是Excle的时间格式就转成文本格式
							tdList.add(changeDateCellToString(cell1));
						}else if(celldataFormat == 20 || celldataFormat == 32){
							tdList.add(changeTimeCellToString(cell1));
						}else if(celldataFormat == 22){
							tdList.add(changeDateTimeCellToString(cell1));
						}else if (cell1.getCellType() == XSSFCell.CELL_TYPE_STRING) {
							tdList.add(cell1.getStringCellValue());
						}else if (cell1.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
							tdList.add(changeNUMCellToString(cell1));
						}else if(cell1.getCellType() == XSSFCell.CELL_TYPE_FORMULA){
							try {  
								tdList.add(cell1.getStringCellValue());  
							} catch (IllegalStateException e) {  
								tdList.add(changeNUMCellToString(cell1));
							}  
						}else{
							cell1.setCellType(Cell.CELL_TYPE_STRING);
							tdList.add(cell1.getStringCellValue());
						}
					} else {
						tdList.add("");
					}
				} else if (fillMaxtdLength) {
					tdList.add("");
				}
			}
			if (removeEmptyRow) {
				if (tdList != null && !tdList.isEmpty()) {
					boolean addFlag = false;
					for (String str : tdList) {
						if (StringUtils.isNotEmpty(str)) {
							addFlag = true;
							break;
						}
					}
					if (addFlag) {
						trList.add(tdList);
					}
				}
			} else {
				trList.add(tdList);
			}
		}
		return trList;
	}
	
	
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
		if(cell.getCellType() != XSSFCell.CELL_TYPE_FORMULA){
			HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
			String cellFormatted = dataFormatter.formatCellValue(cell);
			return StringUtils.contains(cellFormatted, ".");
		}
		return true;
	}
	
	
	
	
	
	
	public static List<List<String>> getExcelAsFile0816解决不同数据转换问题(
//			MultipartFile mulfile,
			String fileUrl, boolean removeEmptyRow,
			boolean fillMaxtdLength) throws IOException, EncryptedDocumentException, InvalidFormatException {
		// 1.得到Excel常用对象
		InputStream fileInputStream = null;
		if (StringUtils.isNotEmpty(fileUrl)) {
			fileInputStream = new FileInputStream(fileUrl);
		} else {
//			fileInputStream = mulfile.getInputStream();
		}

		XSSFWorkbook wb = new XSSFWorkbook(fileInputStream); 
		XSSFSheet sheet = wb.getSheetAt(0);

		int pLength = sheet.getPhysicalNumberOfRows();

		// 总行数
		int trLength = sheet.getLastRowNum();
		trLength = trLength < pLength ? pLength : trLength + 1;
		// 4.得到Excel工作表的行
		XSSFRow row = sheet.getRow(0);
		// 总列数
		int tdLength = row.getLastCellNum();

		int maxtdLength = 0;
		if (fillMaxtdLength) {
			for (int i = 0; i < trLength; i++) {
				if (i == 0 || (maxtdLength < sheet.getRow(i).getLastCellNum())) {
					maxtdLength = sheet.getRow(i).getLastCellNum();
				}
			}
			tdLength = maxtdLength;
		}

		List<List<String>> trList = new ArrayList<List<String>>();
		for (int i = 0; i < trLength; i++) {
			// 得到Excel工作表的行
			XSSFRow row1 = sheet.getRow(i);
			List<String> tdList = new ArrayList<String>();
			for (int j = 0; j < tdLength; j++) {
				if (j < row1.getLastCellNum()) {
					if (row1 == null) {
						break;
					}
					// 得到Excel工作表指定行的单元格
					XSSFCell cell1 = row1.getCell(j);
					if (cell1 != null) {
						// 获得每一列中的值。。。
						short  celldataFormat = cell1.getCellStyle().getDataFormat();
						if (celldataFormat == 14 || celldataFormat == 31 || celldataFormat == 57 || celldataFormat == 58) {// 如果是Excle的时间格式就转成文本格式
							tdList.add(changeDateCellToString(cell1));
						}else if(celldataFormat == 20 || celldataFormat == 32){
							tdList.add(changeTimeCellToString(cell1));
						}else if(celldataFormat == 22){
							tdList.add(changeDateTimeCellToString(cell1));
						}else if (cell1.getCellType() == XSSFCell.CELL_TYPE_STRING) {
							tdList.add(cell1.getStringCellValue());
						}else if (cell1.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
							tdList.add(changeNUMCellToString(cell1));
						}else{
							cell1.setCellType(Cell.CELL_TYPE_STRING);
							tdList.add(cell1.getStringCellValue());
						}
					} else {
						tdList.add("");
					}
				} else if (fillMaxtdLength) {
					tdList.add("");
				}
			}
			if (removeEmptyRow) {
				if (tdList != null && !tdList.isEmpty()) {
					boolean addFlag = false;
					for (String str : tdList) {
						if (StringUtils.isNotEmpty(str)) {
							addFlag = true;
							break;
						}
					}
					if (addFlag) {
						trList.add(tdList);
					}
				}
			} else {
				trList.add(tdList);
			}
		}
		return trList;
	}
	
	
	// @SuppressWarnings("unused")
	public static List<List<String>> getExcelAsFile0731(
//			MultipartFile mulfile, 
			String fileUrl, boolean removeEmptyRow,
			boolean fillMaxtdLength) throws IOException, EncryptedDocumentException, InvalidFormatException {
		// 1.得到Excel常用对象
		InputStream fileInputStream = null;
		if (StringUtils.isNotEmpty(fileUrl)) {
			fileInputStream = new FileInputStream(fileUrl);
		} else {
//			fileInputStream = mulfile.getInputStream();
		}

		Workbook wb = WorkbookFactory.create(fileInputStream);
		Sheet sheet = wb.getSheetAt(0);

		int pLength = sheet.getPhysicalNumberOfRows();

		// 总行数
		int trLength = sheet.getLastRowNum();
		trLength = trLength < pLength ? pLength : trLength + 1;
		// 4.得到Excel工作表的行
		Row row = sheet.getRow(0);
		// 总列数
		int tdLength = row.getLastCellNum();

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
		// Cell cell = row.getCell((short) 1);
		// 6.得到单元格样式
		// CellStyle cellStyle = cell.getCellStyle();
		
		List<List<String>> trList = new ArrayList<List<String>>();
		for (int i = 0; i < trLength; i++) {
			// 得到Excel工作表的行
			Row row1 = sheet.getRow(i);
			List<String> tdList = new ArrayList<String>();
			for (int j = 0; j < tdLength; j++) {
				if (j < row1.getLastCellNum()) {
					if (row1 == null) {
						break;
					}
					// 得到Excel工作表指定行的单元格
					Cell cell1 = row1.getCell(j);
					/**
					 * 为了处理：Excel异常Cannot get a text value from a numeric cell
					 * 将所有列中的内容都设置成String类型格式
					 */
					if (cell1 != null) {
						// 获得每一列中的值。。。
						cell1.setCellType(Cell.CELL_TYPE_STRING);
						tdList.add(cell1.getStringCellValue());
					} else {
						tdList.add("");
					}
				} else if (fillMaxtdLength) {
					tdList.add("");
				}
			}
			if (removeEmptyRow) {
				if (tdList != null && !tdList.isEmpty()) {
					boolean addFlag = false;
					for (String str : tdList) {
						if (StringUtils.isNotEmpty(str)) {
							addFlag = true;
							break;
						}
					}
					if (addFlag) {
						trList.add(tdList);
					}
				}
			} else {
				trList.add(tdList);
			}
		}
		return trList;
	}
	
	
	public static List<List<String>> getExcelAsFile0726_1(
//			MultipartFile mulfile,
			String fileUrl)
			throws IOException, EncryptedDocumentException, InvalidFormatException {
		// 1.得到Excel常用对象
		InputStream fileInputStream = null;
		if (StringUtils.isNotEmpty(fileUrl)) {
			fileInputStream = new FileInputStream(fileUrl);
		} else {
//			fileInputStream = mulfile.getInputStream();
		}

		Workbook wb = WorkbookFactory.create(fileInputStream);
		Sheet sheet = wb.getSheetAt(0);

		int pLength = sheet.getPhysicalNumberOfRows();

		// 总行数
		int trLength = sheet.getLastRowNum();
		trLength = trLength < pLength ? pLength : trLength + 1;
		// 4.得到Excel工作表的行
		Row row = sheet.getRow(0);
		// 总列数
		int tdLength = row.getLastCellNum();

		// 5.得到Excel工作表指定行的单元格
		Cell cell = row.getCell((short) 1);
		// 6.得到单元格样式
		CellStyle cellStyle = cell.getCellStyle();
		List<List<String>> trList = new ArrayList<List<String>>();

		for (int i = 0; i < trLength; i++) {
			// 得到Excel工作表的行
			Row row1 = sheet.getRow(i);
			List<String> tdList = new ArrayList<String>();
			for (int j = 0; j < tdLength; j++) {
				if (row1 == null) {
					break;
				}
				// 得到Excel工作表指定行的单元格
				Cell cell1 = row1.getCell(j);
				/**
				 * 为了处理：Excel异常Cannot get a text value from a numeric cell
				 * 将所有列中的内容都设置成String类型格式
				 */
				if (cell1 != null) {
					// 获得每一列中的值。。。
					cell1.setCellType(Cell.CELL_TYPE_STRING);
					tdList.add(cell1.getStringCellValue());
				} else {
					tdList.add("");
				}
			}
			trList.add(tdList);
		}
		return trList;
	}

	/**
	 * 得到Excel，并解析内容
	 * 
	 * @param fileUrl
	 * @param removeEmptyRow
	 * @throws IOException
	 * @throws InvalidFormatException
	 * @throws EncryptedDocumentException
	 */
	public static List<List<String>> getExcelAsFile0726_2(
//			MultipartFile mulfile, 
			String fileUrl, boolean removeEmptyRow)
			throws IOException, EncryptedDocumentException, InvalidFormatException {
		// 1.得到Excel常用对象
		InputStream fileInputStream = null;
		if (StringUtils.isNotEmpty(fileUrl)) {
			fileInputStream = new FileInputStream(fileUrl);
		} else {
//			fileInputStream = mulfile.getInputStream();
		}

		Workbook wb = WorkbookFactory.create(fileInputStream);
		Sheet sheet = wb.getSheetAt(0);

		int pLength = sheet.getPhysicalNumberOfRows();

		// 总行数
		int trLength = sheet.getLastRowNum();
		trLength = trLength < pLength ? pLength : trLength + 1;
		// 4.得到Excel工作表的行
		Row row = sheet.getRow(0);
		// 总列数
		int tdLength = row.getLastCellNum();

		// 5.得到Excel工作表指定行的单元格
		Cell cell = row.getCell((short) 1);
		// 6.得到单元格样式
		CellStyle cellStyle = cell.getCellStyle();
		List<List<String>> trList = new ArrayList<List<String>>();

		for (int i = 0; i < trLength; i++) {
			// 得到Excel工作表的行
			Row row1 = sheet.getRow(i);
			List<String> tdList = new ArrayList<String>();
			for (int j = 0; j < tdLength; j++) {
				if (row1 == null) {
					break;
				}
				// 得到Excel工作表指定行的单元格
				Cell cell1 = row1.getCell(j);
				/**
				 * 为了处理：Excel异常Cannot get a text value from a numeric cell
				 * 将所有列中的内容都设置成String类型格式
				 */
				if (cell1 != null) {
					// 获得每一列中的值。。。
					cell1.setCellType(Cell.CELL_TYPE_STRING);
					tdList.add(cell1.getStringCellValue());
				} else {
					tdList.add("");
				}
			}
			if (removeEmptyRow) {
				if (tdList != null && !tdList.isEmpty()) {
					boolean addFlag = false;
					for (String str : tdList) {
						if (StringUtils.isNotEmpty(str)) {
							addFlag = true;
							break;
						}
					}
					if (addFlag) {
						trList.add(tdList);
					}
				}
			} else {
				trList.add(tdList);
			}
		}
		return trList;
	}

	
	
	public static List<List<String>> getExcelAsFile0712(
//			MultipartFile mulfile,
			String fileUrl, String userType)
			throws IOException, EncryptedDocumentException, InvalidFormatException {
		// 1.得到Excel常用对象
		InputStream fileInputStream = null;
		if (StringUtils.isNotEmpty(fileUrl)) {
			fileInputStream = new FileInputStream(fileUrl);
		} else {
//			fileInputStream = mulfile.getInputStream();
		}

		Workbook wb = WorkbookFactory.create(fileInputStream);
		Sheet sheet = wb.getSheetAt(0);

		int pLength = sheet.getPhysicalNumberOfRows();
		System.out.println("物理行数 pLength:" + pLength);

		// 总行数
		int trLength = sheet.getLastRowNum();
		System.out.println("总行数：" + trLength);
		trLength = trLength < pLength ? pLength : trLength + 1;
		// 4.得到Excel工作表的行
		Row row = sheet.getRow(0);
		// 总列数
		int tdLength = row.getLastCellNum();
		System.out.println("总列数：" + tdLength);

		if ("1".equals(userType)) {
//			tdLength = 15;
		} else if ("2".equals(userType)) {
//			tdLength = 10;
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
				if (row1 == null) {
					break;
				}
				// 得到Excel工作表指定行的单元格
				Cell cell1 = row1.getCell(j);
				/**
				 * 为了处理：Excel异常Cannot get a text value from a numeric cell
				 * 将所有列中的内容都设置成String类型格式
				 */
				if (cell1 != null) {
					 // 获得每一列中的值。。。
//                         if(HSSFDateUtil.isCellDateFormatted(cell1)){//如果是Excle的时间格式就转成文本格式
//                              Date date = cell1.getDateCellValue();
//                              SimpleDateFormat dateFormate = new SimpleDateFormat("yyyy-MM-dd");
//                              String dateStr = dateFormate.format(date);
//                              tdList.add(dateStr);
//                              System.out.print(dateStr+ "\t\t");
//                         }else{
                              cell1.setCellType(Cell.CELL_TYPE_STRING);
                              tdList.add(cell1.getStringCellValue());
                              System.out.print(cell1.getStringCellValue()+ "\t\t");
//                         }
				} else {
					tdList.add("");
					System.out.print("\t\t");
				}
				// System.out.print(cell1 + "\t\t");
				// System.out.print(cell1.getStringCellValue() + "\t\t");
			}
			trList.add(tdList);
			System.out.println();
		}
		return trList;
	}

	
	
	public static List<List<String>> getExcelAsFile0915解决空白行问题(
//			MultipartFile mulfile,
			String fileUrl, boolean removeEmptyRow,
			boolean fillMaxtdLength) throws IOException, EncryptedDocumentException, InvalidFormatException {
		// 1.得到Excel常用对象
		InputStream fileInputStream = null;
		if (StringUtils.isNotEmpty(fileUrl)) {
			fileInputStream = new FileInputStream(fileUrl);
		} else {
//			fileInputStream = mulfile.getInputStream();
		}

		XSSFWorkbook wb = new XSSFWorkbook(fileInputStream); 
		XSSFSheet sheet = wb.getSheetAt(0);

		int pLength = sheet.getPhysicalNumberOfRows();

		// 总行数
		int trLength = sheet.getLastRowNum();
		trLength = trLength < pLength ? pLength : trLength + 1;
		// 4.得到Excel工作表的行
		XSSFRow row = sheet.getRow(0);
		// 总列数
		int tdLength = row.getLastCellNum();

		int maxtdLength = 0;
		if (fillMaxtdLength) {
			for (int i = 0; i < trLength; i++) {
				if (sheet.getRow(i)!=null && (i == 0 || (maxtdLength < sheet.getRow(i).getLastCellNum()))) {
						maxtdLength = sheet.getRow(i).getLastCellNum();
				}
			}
			tdLength = maxtdLength;
		}

		List<List<String>> trList = new ArrayList<List<String>>();
		for (int i = 0; i < trLength; i++) {
			// 得到Excel工作表的行
			XSSFRow row1 = sheet.getRow(i);
			if (row1 == null) {
				continue;
			}
			List<String> tdList = new ArrayList<String>();
			for (int j = 0; j < tdLength; j++) {
				if (j < row1.getLastCellNum()) {
					// 得到Excel工作表指定行的单元格
					XSSFCell cell1 = row1.getCell(j);
					if (cell1 != null) {
						// 获得每一列中的值。。。
						short  celldataFormat = cell1.getCellStyle().getDataFormat();
						if (celldataFormat == 14 || celldataFormat == 31 || celldataFormat == 57 || celldataFormat == 58) {// 如果是Excle的时间格式就转成文本格式
							tdList.add(changeDateCellToString(cell1));
						}else if(celldataFormat == 20 || celldataFormat == 32){
							tdList.add(changeTimeCellToString(cell1));
						}else if(celldataFormat == 22){
							tdList.add(changeDateTimeCellToString(cell1));
						}else if (cell1.getCellType() == XSSFCell.CELL_TYPE_STRING) {
							tdList.add(cell1.getStringCellValue());
						}else if (cell1.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
							tdList.add(changeNUMCellToString(cell1));
						}else if(cell1.getCellType() == XSSFCell.CELL_TYPE_FORMULA){
							try {  
								tdList.add(cell1.getStringCellValue());  
							} catch (IllegalStateException e) {  
								tdList.add(changeNUMCellToString(cell1));
							}  
						}else{
							cell1.setCellType(Cell.CELL_TYPE_STRING);
							tdList.add(cell1.getStringCellValue());
						}
					} else {
						tdList.add("");
					}
				} else if (fillMaxtdLength) {
					tdList.add("");
				}
			}
			if (removeEmptyRow) {
				if (tdList != null && !tdList.isEmpty()) {
					boolean addFlag = false;
					for (String str : tdList) {
						if (StringUtils.isNotEmpty(str)) {
							addFlag = true;
							break;
						}
					}
					if (addFlag) {
						trList.add(tdList);
					}
				}
			} else {
				trList.add(tdList);
			}
		}
		return trList;
	}
}
