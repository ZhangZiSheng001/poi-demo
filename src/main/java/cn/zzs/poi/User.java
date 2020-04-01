package cn.zzs.poi;

import java.util.Date;

/**
 * <p>用户实体类，与数据库表一一对应</p>
 * @author: zzs
 * @date: 2020年2月25日 上午10:01:21
 */
public class User {

    private String id;

    /**
     * <p>用户名</p>
     */
    private String name;

    /**
     * <p>性别</p>
     */
    private Integer gender = 0;

    /**
     * <p>年龄</p>
     */
    private Integer age;

    /**
     * <p>创建时间</p>
     */
    private Date gmtCreate;

    /**
     * <p>最近修改时间</p>
     */
    private Date gmtModified;

    /**
     * <p>是否删除</p>
     */
    private Integer deleted = 0;

    /**
     * <p>电话号码</p>
     */
    private String phone;

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

    public Date getGmtCreate() {
        return gmtCreate;
    }

    public void setGmtCreate(Date gmtCreate) {
        this.gmtCreate = gmtCreate;
    }

    public Date getGmtModified() {
        return gmtModified;
    }

    public void setGmtModified(Date gmtModified) {
        this.gmtModified = gmtModified;
    }

    public Integer getDeleted() {
        return deleted;
    }

    public void setDeleted(Integer deleted) {
        this.deleted = deleted;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    @Override
    public String toString() {
        StringBuilder builder = new StringBuilder();
        builder.append("User [id=");
        builder.append(id);
        builder.append(", name=");
        builder.append(name);
        builder.append(", gender=");
        builder.append(gender);
        builder.append(", age=");
        builder.append(age);
        builder.append(", gmtCreate=");
        builder.append(gmtCreate);
        builder.append(", gmtModified=");
        builder.append(gmtModified);
        builder.append(", deleted=");
        builder.append(deleted);
        builder.append(", phone=");
        builder.append(phone);
        builder.append("]");
        return builder.toString();
    }
}
