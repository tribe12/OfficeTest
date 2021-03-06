package cn.wgh.excle.XSSF;

import static org.junit.Assert.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;  

public class XSSFTest2 {
	 /** 
     * 得到Excel，并解析内容  对2007及以上版本 使用XSSF解析 
     * @param file 
     * @throws FileNotFoundException 
     * @throws IOException 
     * @throws InvalidFormatException  
     */  
	@Test
    public void getExcelAsFile(String file) throws FileNotFoundException, IOException, InvalidFormatException{  
      //1.得到Excel常用对象  
      POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream("E:/test/xlsTest/HSSF/stuTest2.xls"));  
      //2.得到Excel工作簿对象  
      HSSFWorkbook HSSFwb = new HSSFWorkbook(fs);  
          
        InputStream ins = null;     
        Workbook wb = null;     
            ins=new FileInputStream(new File(file));     
            //ins= ExcelService.class.getClassLoader().getResourceAsStream(filePath);     
            try {
				wb = WorkbookFactory.create(ins);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}     
            ins.close();     
          
          
        //3.得到Excel工作表对象  
        Sheet sheet = wb.getSheetAt(0);  
        //总行数  
        int trLength = sheet.getLastRowNum();  
        //4.得到Excel工作表的行  
        Row row = sheet.getRow(0);  
        //总列数  
        int tdLength = row.getLastCellNum();  
        //5.得到Excel工作表指定行的单元格  
        Cell cell = row.getCell((short)1);  
        //6.得到单元格样式  
        CellStyle cellStyle = cell.getCellStyle();  
  
        for(int i=5;i<trLength;i++){  
            //得到Excel工作表的行  
            Row row1 = sheet.getRow(i);  
            for(int j=0;j<tdLength;j++){  
            //得到Excel工作表指定行的单元格  
            Cell cell1 = row1.getCell(j);  
            /** 
             * 为了处理：Excel异常Cannot get a text value from a numeric cell 
             * 将所有列中的内容都设置成String类型格式 
             */  
            if(cell1!=null){  
                  cell1.setCellType(Cell.CELL_TYPE_STRING);  
             }  
              
            if(j==5&&i<=10){  
                cell1.setCellValue("1000");  
            }  
              
            //获得每一列中的值  
            System.out.print(cell1+"                   ");  
            }  
            System.out.println();  
        }  
          
        //将修改后的数据保存  
        OutputStream out = new FileOutputStream(file);  
                wb.write(out);  
    }  
      
      
    /** 
     * 创建Excel，并写入内容 
     */  
    public static void CreateExcel(){  
          
        //1.创建Excel工作薄对象  
        HSSFWorkbook wb = new HSSFWorkbook();  
        //2.创建Excel工作表对象       
        HSSFSheet sheet = wb.createSheet("new Sheet");  
        //3.创建Excel工作表的行     
        HSSFRow row = sheet.createRow(6);  
        //4.创建单元格样式  
        CellStyle cellStyle =wb.createCellStyle();  
          // 设置这些样式  
        cellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);  
        cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);  
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);  
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);  
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);  
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);  
            
            
            
        //5.创建Excel工作表指定行的单元格  
        row.createCell(0).setCellStyle(cellStyle);  
        //6.设置Excel工作表的值  
        row.createCell(0).setCellValue("aaaa");  
          
        row.createCell(1).setCellStyle(cellStyle);  
        row.createCell(1).setCellValue("bbbb");  
          
          
        //设置sheet名称和单元格内容  
        wb.setSheetName(0,"第一张工作表");  
        //设置单元格内容   cell.setCellValue("单元格内容");  
          
        // 最后一步，将文件存到指定位置  
                try  
                {  
                    FileOutputStream fout = new FileOutputStream("E:/students.xls");  
                    wb.write(fout);  
                    fout.close();  
                }  
                catch (Exception e)  
                {  
                    e.printStackTrace();  
                }  
    }  
      
    /** 
     * 创建Excel的实例 
     * @throws ParseException  
     */  
    public static void CreateExcelDemo1() throws ParseException{  
        List list = new ArrayList();  
        SimpleDateFormat df = new SimpleDateFormat("yyyy-mm-dd");  
        Student user1 = new Student(1, "张三", 16,true, df.parse("1997-03-12"));  
        Student user2 = new Student(2, "李四", 17,true, df.parse("1996-08-12"));  
        Student user3 = new Student(3, "王五", 26,false, df.parse("1985-11-12"));  
        list.add(user1);  
        list.add(user2);  
        list.add(user3);  
          
          
        // 第一步，创建一个webbook，对应一个Excel文件  
                HSSFWorkbook wb = new HSSFWorkbook();  
                // 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet  
                HSSFSheet sheet = wb.createSheet("学生表一");  
                // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short  
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
  
                // 第五步，写入实体数据 实际应用中这些数据从数据库得到，  
  
                for (int i = 0; i < list.size(); i++)  
                {  
                    row = sheet.createRow((int) i + 1);  
                    Student stu = (Student) list.get(i);  
                    // 第四步，创建单元格，并设置值  
                    row.createCell((short) 0).setCellValue((double) stu.getId());  
                    row.createCell((short) 1).setCellValue(stu.getName());  
                    row.createCell((short) 2).setCellValue((double) stu.getAge());  
                    row.createCell((short)3).setCellValue(stu.isSex()==true?"男":"女");  
                    cell = row.createCell((short) 4);  
                    cell.setCellValue(new SimpleDateFormat("yyyy-mm-dd").format(stu  
                            .getBirthday()));  
                }  
                // 第六步，将文件存到指定位置  
                try  
                {  
                    FileOutputStream fout = new FileOutputStream("E:/students.xls");  
                    wb.write(fout);  
                    fout.close();  
                }  
                catch (Exception e)  
                {  
                    e.printStackTrace();  
                }  
          
          
          
    }  
}
