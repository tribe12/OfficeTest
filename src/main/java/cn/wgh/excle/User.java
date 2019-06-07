package cn.wgh.excle;

public class User {
	private String id;
	private String name;
	private String sex;
	private String age;
	private String chengji;
	private String xueyear;
	private String yuwen;
	private String shuxue;
	private String yingyu;
	private String pingyu;
	private String zongjie;

	User(String id, String name, String sex, String age, String xueyear, String yuwen, String shuxue, String yingyu,
			String zongjie) {
		this.id = id;
		this.name = name;
		this.sex = sex;
		this.age = age;
		this.xueyear = xueyear;
		this.yuwen = yuwen;
		this.shuxue = shuxue;
		this.yingyu = yingyu;
		this.zongjie = zongjie;
	}

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getSex() {
		return sex;
	}

	public void setSex(String sex) {
		this.sex = sex;
	}

	public String getAge() {
		return age;
	}

	public void setAge(String age) {
		this.age = age;
	}

	public String getChengji() {
		return chengji;
	}

	public void setChengji(String chengji) {
		this.chengji = chengji;
	}

	public String getXueyear() {
		return xueyear;
	}

	public void setXueyear(String xueyear) {
		this.xueyear = xueyear;
	}

	public String getYuwen() {
		return yuwen;
	}

	public void setYuwen(String yuwen) {
		this.yuwen = yuwen;
	}

	public String getShuxue() {
		return shuxue;
	}

	public void setShuxue(String shuxue) {
		this.shuxue = shuxue;
	}

	public String getYingyu() {
		return yingyu;
	}

	public void setYingyu(String yingyu) {
		this.yingyu = yingyu;
	}

	public String getPingyu() {
		return pingyu;
	}

	public void setPingyu(String pingyu) {
		this.pingyu = pingyu;
	}

	public String getZongjie() {
		return zongjie;
	}

	public void setZongjie(String zongjie) {
		this.zongjie = zongjie;
	}

}
