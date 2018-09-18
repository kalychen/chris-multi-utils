package com.chris.multi.model;

import com.chris.multi.poi.xls.XlsColumn;
import com.chris.multi.poi.xls.XlsSheet;

/**
 * Created by Chris Chen
 * 2018/09/17
 * Explain:
 */
@XlsSheet("用户表")
public class UserModel {
    @XlsColumn("编号")
    public int id;
    @XlsColumn("姓名")
    public String name;
    @XlsColumn("年龄")
    public int age;
    @XlsColumn("地址")
    public String address;

    public UserModel() {
    }

    public UserModel(int id, String name, int age, String address) {
        this.id = id;
        this.name = name;
        this.age = age;
        this.address = address;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }
}
