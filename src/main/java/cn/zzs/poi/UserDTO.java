package cn.zzs.poi;

import java.io.Serializable;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.ContentRowHeight;

/**
 * <p>用户 DTO类</p>
 * @author: zzs
 * @date: 2020年2月25日 上午10:12:21
 */
@ContentRowHeight(16)
public class UserDTO implements Serializable {

	private static final long serialVersionUID = 1L;
	
	@ExcelIgnore
	private String id;

	/**
	 * <p>用户名</p>
	 */
	@ExcelProperty(value = { "用户名" }, index = 1)
	private String name;

	/**
	 * <p>性别</p>
	 */
	@ExcelProperty(value = { "性别" }, index = 2)
	private String genderStr;

	/**
	 * <p>年龄</p>
	 */
	@ExcelProperty(value = { "年龄" }, index = 3)
	private Integer age;

	/**
	 * <p>电话号码</p>
	 */
	@ExcelProperty(value = { "手机号" }, index = 4)
	@ColumnWidth(14)
	private String phone;
	
	@ExcelIgnore
	private Integer gender = 0;

	public UserDTO() {
		super();
	}

	public UserDTO(User user) {
		super();
		setId(user.getId());
		setName(user.getName());
		setAge(user.getAge());
		setGender(user.getGender());
		setPhone(user.getPhone());
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

	public Integer getGender() {
		return gender;
	}

	public void setGender(Integer gender) {
		this.gender = gender;
	}

	public Integer getAge() {
		return age;
	}

	public void setAge(Integer age) {
		this.age = age;
	}

	public String getPhone() {
		return phone;
	}

	public void setPhone(String phone) {
		this.phone = phone;
	}

	public String getGenderStr() {
		if(genderStr == null) {
			if(getGender() == 0) {
				genderStr = "男";
			} else {
				genderStr = "女";
			}
		}
		return genderStr;
	}

	public void setGenderStr(String genderStr) {
		if("女".equals(genderStr)) {
			setGender(1);
		}
		this.genderStr = genderStr;
	}

	@Override
	public String toString() {
		StringBuilder builder = new StringBuilder();
		builder.append("UserDTO [id=");
		builder.append(id);
		builder.append(", name=");
		builder.append(name);
		builder.append(", genderStr=");
		builder.append(genderStr);
		builder.append(", age=");
		builder.append(age);
		builder.append(", phone=");
		builder.append(phone);
		builder.append(", gender=");
		builder.append(gender);
		builder.append("]");
		return builder.toString();
	}
	
}
