package cn.wgh.excle.utils;

import java.util.List;

import lombok.Data;

/**
 * 表头和表内容
 * @author guohui.wang
 */

@Data
public class TitleAndValForTable {
	private List<String> title ;
	private List<List<String>> valueList;
}
