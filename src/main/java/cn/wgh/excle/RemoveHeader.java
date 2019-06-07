package cn.wgh.excle;

import java.util.List;

public class RemoveHeader {
	private int sort;
	private int emptyCount;
	private List<String> empHeaderList;

	public RemoveHeader() {

	}

	public RemoveHeader(int sort, int emptyCount, List<String> empHeaderList) {
		this.sort = sort;
		this.emptyCount = emptyCount;
		this.empHeaderList = empHeaderList;
	}

	public int getSort() {
		return sort;
	}

	public void setSort(int sort) {
		this.sort = sort;
	}

	public int getEmptyCount() {
		return emptyCount;
	}

	public void setEmptyCount(int emptyCount) {
		this.emptyCount = emptyCount;
	}

	public List<String> getEmpHeaderList() {
		return empHeaderList;
	}

	public void setEmpHeaderList(List<String> empHeaderList) {
		this.empHeaderList = empHeaderList;
	}
}
