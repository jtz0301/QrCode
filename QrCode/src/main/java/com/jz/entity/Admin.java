package com.jz.entity;

public class Admin {
	
	private Integer id;
	private String phone;
	private String name;
	public Admin() {
		super();
		// TODO Auto-generated constructor stub
	}
	public Admin(Integer id, String phone, String name) {
		super();
		this.id = id;
		this.phone = phone;
		this.name = name;
	}
	public Integer getId() {
		return id;
	}
	public void setId(Integer id) {
		this.id = id;
	}
	public String getPhone() {
		return phone;
	}
	public void setPhone(String phone) {
		this.phone = phone;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	@Override
	public String toString() {
		return "Admin [id=" + id + ", phone=" + phone + ", name=" + name + "]";
	}
	
	

}
