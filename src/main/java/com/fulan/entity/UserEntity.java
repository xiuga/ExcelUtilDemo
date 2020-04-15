package com.fulan.entity;

import com.fulan.annotation.ExcelIO;

/**
 * @author xiang
 * @date 2020年2月16日 
 */
public class UserEntity {
	@ExcelIO(column = "A", name = "姓名")
	private String name;
	@ExcelIO(column = "B", name = "年龄")
	private Integer age;
	@ExcelIO(column = "C", name = "电话号码")
	private String phoneNum;
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public Integer getAge() {
		return age;
	}
	public void setAge(Integer age) {
		this.age = age;
	}
	public String getPhoneNum() {
		return phoneNum;
	}
	public void setPhoneNum(String phoneNum) {
		this.phoneNum = phoneNum;
	}
}
