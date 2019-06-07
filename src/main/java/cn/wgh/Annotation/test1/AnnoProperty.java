package cn.wgh.Annotation.test1;

public class AnnoProperty implements Comparable<AnnoProperty>{
	private String name;
	private Integer sort;
	private String property;
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public Integer getSort() {
		return sort;
	}
	public void setSort(Integer sort) {
		this.sort = sort;
	}
	public String getProperty() {
		return property;
	}
	public void setProperty(String property) {
		this.property = property;
	}
	@Override
	public String toString() {
		return "AnnoProperty [name=" + name + ", sort=" + sort + ", property=" + property + "]";
	}

//	@Override
	public int compareTo(AnnoProperty o) {
		return this.getSort().compareTo(o.getSort());
	}
	
}
