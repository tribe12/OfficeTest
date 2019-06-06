package cn.wgh.excle;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.IndexedColors;

import com.alibaba.excel.metadata.Font;
import com.alibaba.excel.metadata.TableStyle;

public class DataUtil {
	public static TableStyle creatTableStyle() {
		TableStyle tableStyle = new TableStyle();
		
		/*表头*/
		Font headFont = new Font();
		headFont.setBold(true);//字体加粗
		headFont.setFontHeightInPoints((short) 12);
		headFont.setFontName("楷体");
		tableStyle.setTableHeadFont(headFont);
		//表头背景色
		tableStyle.setTableHeadBackGroundColor(IndexedColors.GREEN);
	
		/*表格主体*/
		Font contentfont = new Font();
		contentfont.setBold(true);
		contentfont.setFontHeightInPoints((short) 12);
		contentfont.setFontName("黑体");
		tableStyle.setTableContentFont(contentfont);
		tableStyle.setTableContentBackGroundColor(IndexedColors.YELLOW);
		return tableStyle;
	}
	
	
    public static List<List<String>> createTestListStringHead(){
        // 模型上没有注解，表头数据动态传入
        List<List<String>> head = new ArrayList<List<String>>();
        List<String> headCoulumn1 = new ArrayList<String>();
        List<String> headCoulumn2 = new ArrayList<String>();
        List<String> headCoulumn3 = new ArrayList<String>();
        List<String> headCoulumn4 = new ArrayList<String>();
        List<String> headCoulumn5 = new ArrayList<String>();

        headCoulumn1.add("第一列");headCoulumn1.add("第一列");headCoulumn1.add("第一列");
        headCoulumn2.add("第一列");headCoulumn2.add("第一列");headCoulumn2.add("第一列");

        headCoulumn3.add("第二列");headCoulumn3.add("第二列");headCoulumn3.add("第二列");
        
        headCoulumn4.add("第三列");headCoulumn4.add("第三列2");headCoulumn4.add("第三列2");
        
        headCoulumn5.add("第一列");headCoulumn5.add("第3列");headCoulumn5.add("第4列");

        head.add(headCoulumn1);
        head.add(headCoulumn2);
        head.add(headCoulumn3);
        head.add(headCoulumn4);
        head.add(headCoulumn5);
        return head;
    }
}
