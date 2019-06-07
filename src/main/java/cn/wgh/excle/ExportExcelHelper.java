package cn.wgh.excle;

import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.IOUtils;

import net.sf.json.JSONArray;

public class ExportExcelHelper<T> {

    private List<Map<String, String>> columnMapping = new ArrayList<Map<String, String>>();

    private String title;

    private List<T> items;

    public ExportExcelHelper(String title,List<T> items){
        this.title = title;
        this.items = items;
    }

    public ExportExcelHelper<T> setColumnMapping(String text,String property){
        Map<String, String> mapping = new HashMap<String, String>();
        mapping.put(text, property);
        this.columnMapping.add(mapping);
        return this;
    }

    public ExportExcelHelper<T> setColumnMapping(List<Map<String, String>> columnMapping){
        this.columnMapping.addAll(columnMapping);
        return this;
    }

    public void executeExport(OutputStream out){
        try{
            int leng = this.columnMapping.size();
            String [] headers = new String [leng];
            String [] cols = new String[leng];
            int i = 0;
            for(Map<String, String> mapping : this.columnMapping){
                headers [i] =  mapping.keySet().iterator().next();
                cols [i] =  mapping.values().iterator().next();
                i ++;
            }
            ExportExcel.exportExcel(title, headers, cols, JSONArray.fromObject(items), null, out);
            out.flush();
        }catch(Exception e){
            throw new RuntimeException(e);
        }finally{
            IOUtils.closeQuietly(out);
        }

    }
}
