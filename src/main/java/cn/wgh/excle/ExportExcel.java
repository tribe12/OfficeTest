package cn.wgh.excle;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
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

public class ExportExcel {


    public static void exportExcel(String title, String[] headers,
            String[] cols, JSONArray dataset,
            Map<String, Map<String, String>> constant, OutputStream out) {

        HSSFWorkbook workbook = new HSSFWorkbook();// 声明一个工作薄
        HSSFSheet sheet = workbook.createSheet(title);// 生成一个表格
        HSSFCellStyle style = workbook.createCellStyle();// 生成一个样式
        HSSFCellStyle style2 = workbook.createCellStyle();// 生成并设置另一个样式
        setHssFWorkStyle(workbook, sheet, style, style2);
        try {
            // 产生表格标题行
            HSSFRow row = sheet.createRow(0);
            HSSFCell cell = null;
            for (int i = 0; i < headers.length; i++) {
                cell = row.createCell(i);
                cell.setCellStyle(style);
                HSSFRichTextString text = new HSSFRichTextString(headers[i]);
                cell.setCellValue(text);
            }
            JSONObject map = null;
            int index = 0;
            if (dataset != null) {
                for (int j = 0; j < dataset.size(); j++) {
                    index++;
                    row = sheet.createRow(index);
                    map = dataset.getJSONObject(j);

                    for (int k = 0; k < cols.length; k++) {
                        cell = row.createCell(k);
                        cell.setCellStyle(style2);
                        String text = map.getString(cols[k]);
                        if (StringUtils.isNotEmpty(text)) {
                            if (constant!=null&&constant.containsKey(cols[k])
                                    && constant.get(cols[k]).containsKey(text)) {
                                cell.setCellValue(constant.get(cols[k]).get(
                                        text));
                            } else
                                cell.setCellValue(text);
                        }
                    }
                }
            }

            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public static void downFile(OutputStream os,InputStream stream){
        int len;
        byte[] bs = new byte[1024];
        try {
            while ((len = stream.read(bs)) != -1) {
                os.write(bs, 0, len);
            }
            os.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }finally{
            IOUtils.closeQuietly(stream);
            IOUtils.closeQuietly(os);
        }
    }

    public static String download(String urlString) throws Exception {
        // 构造URL
        URL url = new URL(urlString);
        // 打开连接
        URLConnection con = url.openConnection();
        // 设置请求超时为5s
        con.setConnectTimeout(5 * 1000);
        // 输入流
        InputStream is = con.getInputStream();

        // 1K的数据缓冲
        byte[] bs = new byte[1024];
        // 读取到的数据长度
        int len;
        // 输出的文件流
        String path = ExportExcel.class.getResource("/").getPath();
        String fileName = path.substring(0, path.indexOf("WEB-INF")) + "tmp"
                + System.getProperty("file.separator") + "qrcode"
                + new Date().getTime() + ".jpg";
        OutputStream os = new FileOutputStream(fileName);
        // 开始读取
        while ((len = is.read(bs)) != -1) {
            os.write(bs, 0, len);
        }
        // 完毕，关闭所有链接
        os.close();
        is.close();
        return fileName;
    }

    // 自定义的方法,插入某个图片到指定索引的位置
    private static void insertImage(HSSFWorkbook wb, HSSFPatriarch pa,
            byte[] data, int row, int column, int index) {
        int x1 = index * 250;
        int y1 = 0;
        int x2 = 255 * 4;
        int y2 = 255;
        HSSFClientAnchor anchor = new HSSFClientAnchor(x1, y1, x2, y2,
                (short) column, row, (short) column, row);
        anchor.setAnchorType(2);
        pa.createPicture(anchor,
                wb.addPicture(data, HSSFWorkbook.PICTURE_TYPE_JPEG));
    }

    // 从图片里面得到字节数组
    private static byte[] getImageData(BufferedImage bi) {
        try {
            ByteArrayOutputStream bout = new ByteArrayOutputStream();
            ImageIO.write(bi, "PNG", bout);
            return bout.toByteArray();
        } catch (Exception exe) {
            exe.printStackTrace();
            return null;
        }
    }

    public static void setHssFWorkStyle(HSSFWorkbook workbook, HSSFSheet sheet,
            HSSFCellStyle style, HSSFCellStyle style2) {
        // 设置表格默认列宽度为15个字节
        sheet.setDefaultColumnWidth(15);
        // 设置这些样式
        style.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 生成一个字体
        HSSFFont font = workbook.createFont();
        font.setColor(HSSFColor.VIOLET.index);
        font.setFontHeightInPoints((short) 12);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        // 把字体应用到当前的样式
        style.setFont(font);

        style2.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
        style2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style2.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style2.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style2.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        // 生成另一个字体
        HSSFFont font2 = workbook.createFont();
        font2.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        // 把字体应用到当前的样式
        style2.setFont(font2);

        // 声明一个画图的顶级管理器
        // HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
        // 定义注释的大小和位置,详见文档
        // HSSFComment comment = patriarch.createComment(new HSSFClientAnchor(0,
        // 0, 0, 0, (short) 4, 2, (short) 6, 5));
        // 设置注释内容
        // comment.setString(new HSSFRichTextString("可以在POI中添加注释！"));
        // 设置注释作者，当鼠标移动到单元格上是可以在状态栏中看到该内容.
        // comment.setAuthor("leno");
    }
}