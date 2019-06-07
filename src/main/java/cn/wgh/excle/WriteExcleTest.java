package cn.wgh.excle;
import static org.junit.Assert.*;

import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Test;
public class WriteExcleTest {
	private HSSFWorkbook workbook = null;
	/**
	 * 
	 * @param response 下载请求的response
	 */
	
	
	@Test
	public void createExcel(HttpServletResponse response){
		
    	//创建workbook
    	workbook = new HSSFWorkbook();
    	//添加Worksheet（不添加sheet时生成的xls文件打开时会报错)
    	Sheet sheet1 = workbook.createSheet("sheet1");  
    	OutputStream out = null;
    	try {	 
    		out = response.getOutputStream();
    		String fileName = "test.xls";// 文件名
    		response.setContentType("application/x-msdownload");
    		response.setHeader("Content-Disposition", "attachment; filename="+ URLEncoder.encode(fileName, "UTF-8"));
    		Row row = workbook.getSheet("sheet1").createRow(0);    //创建第一行  
	    	for(int i = 0;i < 10;i++){
	    		Cell cell = row.createCell(i);
	    		cell.setCellValue("测试数据"+i);
	    	}	
			workbook.write(out);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {  
		    try {  	
		        out.close();  
		    } catch (IOException e) {  
		        e.printStackTrace();
		    }  
		}  
    }
}
