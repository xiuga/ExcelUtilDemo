package com.fulan.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;

import com.fulan.entity.UserEntity;
import com.fulan.utils.ExcelIOUtil;

public class Test1 {
	/*
	@Test
	public void writePoi() throws IOException {
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("sheet1");
		for (int i = 0; i < 10; i++) {
			HSSFRow row = sheet.createRow(i);
			for (int j = 0; j < 10; j++) {
				HSSFCell cell = row.createCell(j);
				cell.setCellValue(i);
			}
		}
		FileOutputStream file = new FileOutputStream("F:\\hello.xls");
		workbook.write(file);
		System.out.println("创建Excel成功！");
	}

	@Test
	public void readPoi() throws IOException {
		FileInputStream inputStream = new FileInputStream("F:\\hello.xls");
		HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
		HSSFSheet sheet = workbook.getSheet("sheet1");
		HSSFRow row = sheet.getRow(0);
		HSSFCell cell = row.getCell(3);
		int numericCellValue = (int) cell.getNumericCellValue();
		System.out.println("读取Excel的值为:" + numericCellValue);
	}
	*/
	
	/*
	 * 测试excel文件流转成java泛型对象。
	 */
	@Test
	public void testUtilImport() {
		FileInputStream inputStream = null;
		try {
			inputStream = new FileInputStream("F:\\export.xls");
			ExcelIOUtil<UserEntity> excelIOUtil = new ExcelIOUtil<UserEntity>(UserEntity.class);// 创建工具类.
			List<UserEntity> list = excelIOUtil.ExcelImport("sheet1", inputStream);
			for (UserEntity userEntity : list) {
				String name = userEntity.getName();
				Integer age = userEntity.getAge();
				String phoneNum = userEntity.getPhoneNum();
				System.out.println("姓名：" + name + "," + "年龄：" + age + "," + "电话号码：" + phoneNum);
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}

	/*
	 * 测试java泛型对象转excel文件流。
	 */
	@Test
	public void testUtilExport() {
		OutputStream outputStream = null;
		try {
			outputStream = new FileOutputStream("F:\\export.xls");
			ExcelIOUtil<UserEntity> excelIOUtil = new ExcelIOUtil<UserEntity>(UserEntity.class);// 创建工具类.
			List<UserEntity> list = new ArrayList<UserEntity>();
			UserEntity entity = null;
			for (int i = 0; i < 10; i++) {
				entity = new UserEntity();
				entity.setName("苹果" + i);
				entity.setAge(i);
				entity.setPhoneNum("6665554440" + i);
				list.add(entity);
			}
			excelIOUtil.ExcelExport(list, "水果", 65536, outputStream);
			System.out.println("导出Excel成功！");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
}
