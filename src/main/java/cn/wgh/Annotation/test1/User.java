package cn.wgh.Annotation.test1;

//implements Comparable<User> 用于排序：http://blog.csdn.net/aitangyong/article/details/54880228

public class User implements Comparable<User>{
	private String id;
	@ParamMapping(name = "年龄", sort = 2)
	private String age;
	@ParamMapping(name = "姓名", sort = 1)
	private String name;

	public User(String id, String name, String age) {
		this.id = id;
		this.name = name;
		this.age = age;
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

	public String getAge() {
		return age;
	}

	public void setAge(String age) {
		this.age = age;
	}

//	@Override
	public int compareTo(User o) {
		return 0;
	}

}
