package cn.wgh.excle;
import java.io.UnsupportedEncodingException;
import java.util.HashMap;
import java.util.Map;
import javax.servlet.http.HttpServletRequest;
import net.sf.json.JSONObject;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang.StringUtils;

/**
 * 构建Request参数
 */
public class Parameter {

    private Map<String, Object> param = new HashMap<String, Object>();

    private HttpServletRequest request;

    private Parameter(HttpServletRequest request) {
        this.request = request;
        this.loadParam(null);
    }

    private Parameter(HttpServletRequest request,String charset) {
        this.request = request;
        this.loadParam(charset);
    }

    private void loadParam(String charset) {
        if (request == null) return;
        @SuppressWarnings("unchecked")
        Map<String, String[]> parameterMap = request.getParameterMap();
        for ( Map.Entry<String, String[]> entry : parameterMap.entrySet() ) {
            String value = entry.getValue()[0];
            this.param.put(entry.getKey(), setCharSet(value,charset));
        }
    }

    private String setCharSet(String value, String charset) {
        if(StringUtils.isBlank(charset) || StringUtils.isBlank(value)) return value;
        try {
            return new String(value.getBytes("ISO-8859-1"),charset);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        return value;
    }

    public Parameter addParam(String name,String value){
        this.param.put(name, value);
        return this;
    }

    public Map<String, Object> getParamMap(){
        return this.param;
    }


    public Integer getInt(String name, Integer defaultValue) {
        return MapUtils.getInteger(param, name, defaultValue);
    }

    public Integer getInt(String name){
        return MapUtils.getInteger(param, name);
    }

    public Long getLong(String name, Long defaultValue) {
        return MapUtils.getLong(param, name, defaultValue);
    }

    public Long getLong(String name) {
        return MapUtils.getLong(param, name);
    }

    public Float getFloat(String name, Float defaultValue) {
        return MapUtils.getFloat(param, name, defaultValue);
    }

    public Float getFloat(String name) {
        return MapUtils.getFloat(param, name);
    }

    public Double getDouble(String name, Double defaultValue) {
        return MapUtils.getDouble(param, name, defaultValue);
    }

    public Double getDouble(String name) {
        return MapUtils.getDouble(param, name);
    }

    public boolean getBooleanValue(String name, boolean defaultValue) {
        return MapUtils.getBooleanValue(param, name, defaultValue);
    }

    public boolean getBooleanValue(String name) {
        return MapUtils.getBooleanValue(param, name);
    }


    public short getShortValue(String name, short defaultValue) {
        return MapUtils.getShortValue(param, name, defaultValue);
    }

    public short getShortValue(String name) {
        return MapUtils.getShortValue(param, name);
    }

    public String getString(String name, String defaultValue) {
        return MapUtils.getString(param, name, defaultValue);
    }

    public String getString(String name) {
        return MapUtils.getString(param, name);
    }

    public boolean hasParam(String name){
        return (this.param.keySet().contains(name));
    }

    public boolean valueIsNotBlank(String name){
        return StringUtils.isNotBlank(getString(name, ""));
    }

    public boolean valueIsNotNull(String name){
        return (this.param.get(name) != null);
    }

    public JSONObject toJson(){
        return JSONObject.fromObject(this.getParamMap());
    }

    static public Parameter createFromRequest(HttpServletRequest request){
        return new Parameter(request);
    }

    static public Parameter createFromRequest(HttpServletRequest request,String charset){
        return new Parameter(request, charset);
    }

    public HttpServletRequest getRequest() {
        return request;
    }
}
