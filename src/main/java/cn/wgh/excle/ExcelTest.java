package cn.wgh.excle;
import static org.junit.Assert.*;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.net.URLConnection;
import java.util.Date;
import java.util.Map;

import javax.imageio.ImageIO;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.junit.Test;
public class ExcelTest {
	
	
	@Test
	public void create() throws Exception {
	    
        HSSFWorkbook wb = new HSSFWorkbook();//创建工作薄  
          
        HSSFFont font = wb.createFont();  
        font.setFontHeightInPoints((short)24);  
        font.setFontName("宋体");  
        font.setColor(HSSFColor.BLACK.index);  
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);  
          
        HSSFCellStyle style = wb.createCellStyle();  
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);  
        style.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);  
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);  
        style.setBorderBottom(HSSFCellStyle.BORDER_THICK);  
        style.setFont(font);  
          
        HSSFSheet sheet = wb.createSheet("test");//创建工作表，名称为test  
          
        int iRow = 0;//行号  
        int iMaxCol = 17;//最大列数  
        HSSFRow row = sheet.createRow(iRow);  
        HSSFCell cell = row.createCell((short)0);  
        cell.setCellValue(new HSSFRichTextString("测试excel"));  
        cell.setCellStyle(style);  
//        sheet.addMergedRegion(new Region(iRow,(short)0,iRow,(short)(iMaxCol-1)));  
          
        ByteArrayOutputStream os = new ByteArrayOutputStream();  
          
        try{  
            wb.write(os);  
        }catch(IOException e){  
            e.printStackTrace();  
            //return null;  
        }  
          
        byte[] xls = os.toByteArray();  
          
        File file = new File("test01.xls");  
        OutputStream out = null;  
        try {  
             out = new FileOutputStream(file);  
             try {  
                out.write(xls);  
            } catch (IOException e) {  
                // TODO Auto-generated catch block  
                e.printStackTrace();  
            }  
        } catch (FileNotFoundException e1) {  
            // TODO Auto-generated catch block  
            e1.printStackTrace();  
        }  
          
	}
}
