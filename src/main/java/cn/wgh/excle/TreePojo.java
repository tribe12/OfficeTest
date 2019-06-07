package cn.wgh.excle;

public class TreePojo {
	private String id;
	private String pid;
	private String name;
	private String code;

	TreePojo() {
	}

	public TreePojo(String id, String pid, String name, String code) {
		this.id = id;
		this.pid = pid;
		this.name = name;
		this.code = code;
	}

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	public String getPid() {
		return pid;
	}

	public void setPid(String pid) {
		this.pid = pid;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getCode() {
		return code;
	}

	public void setCode(String code) {
		this.code = code;
	}

	@Override
	public String toString() {
		return "TreePojo [id=" + id + ", pid=" + pid + ", name=" + name + ", code=" + code + "]";
	}
}
