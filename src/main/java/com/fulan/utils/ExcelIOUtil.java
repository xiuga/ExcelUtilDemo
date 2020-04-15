package com.fulan.utils;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.T;

import com.fulan.annotation.ExcelIO;
import com.fulan.exception.MyException;

/*
 * Excel工具类，里面包含有两个方法。
 * 1：java泛型对象转excel文件流。
 * 2：excel文件流转成java泛型对象。
 * 3：做异常提示,如第几行第几列数据格式错误或者不能为空。
 * 4：时间类型的可以指定转换格式。
  *  使用的技术点：
  *  自定义注解、异常处理、java IO、泛型集合、反射。
  *  提示:在entity类上和属性中加上自定义注解，指定导出的head，每列的中文名，以及顺序，和数据格式规范!    
  *  导入指定从第几行开始，是否允许为空，支持其他正则的验证！
  *  
 * @author xiang
 * @date 2020年2月16日
 */
@SuppressWarnings("hiding")
public class ExcelIOUtil<T> {
	Class<T> clazz;

	/**
	 * @param clazz 泛型对象的反射
	 */
	public ExcelIOUtil(Class<T> clazz) {// 反射生成一个泛型对象
		super();
		this.clazz = clazz;
	}

	/**
	 * 将excel文件流转成java泛型对象
	 *
	 * @param sheetName   工作表的名称
	 * @param sheetSize   每个sheet中数据的行数,此数值必须小于65536
	 * @param inputStream java输入流
	 */
	@SuppressWarnings({ "deprecation", "resource" })
	public List<T> ExcelImport(String sheetName, InputStream inputStream) {
		List<T> list = new ArrayList<T>();
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(inputStream);// 创建一个Excel工作簿
			HSSFSheet sheet = workbook.getSheet(sheetName);
			if (!sheetName.trim().equals("")) {
				sheet = workbook.getSheet(sheetName);// 如果指定sheet名,则取指定sheet中的内容
			}
			if (sheet == null) {
				sheet = workbook.getSheetAt(0); // 如果传入的sheet名不存在则默认指向第1个sheet
			}
			int rows = sheet.getPhysicalNumberOfRows(); // 获取行号
			if (rows > 0) { // 如果excel中有数据
				Field[] fields = clazz.getDeclaredFields();// 通过反射获得类中所有声明的字段
				Map<Integer, Field> fieldsMap = new HashMap<Integer, Field>();// 存放列的序号和field对象数组
				for (Field field : fields) {
					if (field.isAnnotationPresent(ExcelIO.class)) {// 判断是否包含自定义注解ExcelIO
						ExcelIO excelIOAttr = field.getAnnotation(ExcelIO.class);// 返回注解
						int column = getExcelCol(excelIOAttr.column()); // 获取列号
						field.setAccessible(true); // 设置类的私有字段属性可访问.
						fieldsMap.put(column, field);
					} else {
						throw new MyException();
					}
				}
				for (int i = 1; i < rows; i++) { // 从第二行开始取数据(第一行为表头)
					HSSFRow row = sheet.getRow(i);
					int cells = row.getPhysicalNumberOfCells();
					T entity = null; // 定义该泛型类型的对象
					for (int j = 0; j < cells; j++) { // 从第一列开始取数据
						HSSFCell cell = row.getCell(j);
						if (cell != null) {
							String value;
							switch (cell.getCellType()) {
							case HSSFCell.CELL_TYPE_NUMERIC: // 数字
								if (HSSFDateUtil.isCellDateFormatted(cell)) { // 如果为时间格式的内容,需要支持时间类型可以指定转换格式.
									SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
									value = sdf
											.format(HSSFDateUtil.getJavaDate(new Double(cell.getNumericCellValue())));// 数字改为日期字符串
									break;
								} else {
									value = new DecimalFormat("0").format(cell.getNumericCellValue());
								}
								break;
							case HSSFCell.CELL_TYPE_STRING: // 字符串
								value = cell.getStringCellValue();
								break;
							case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean
								value = cell.getBooleanCellValue() + "";
								break;
							case HSSFCell.CELL_TYPE_FORMULA: // 公式
								value = cell.getCellFormula() + "";
								break;
							case HSSFCell.CELL_TYPE_BLANK: // 空值
								value = null;
								break;
							case HSSFCell.CELL_TYPE_ERROR: // 错误
								throw new MyException("第" + (i + 1) + "行第" + (j + 1) + "列数据格式错误!");
							default:
								throw new MyException("第" + (i + 1) + "行第" + (j + 1) + "列数据格式错误!");
							}
							if (value == null) {
								throw new MyException("第" + (i + 1) + "行第" + (j + 1) + "列数据不能为空!");// 数据为空做出提示
							}
							entity = (entity == null ? clazz.newInstance() : entity);
							Field field = fieldsMap.get(j);// 获取对应的field
							if (field == null) {
								continue;
							}
							Class<?> fieldType = field.getType();// 获取此Field对象所表示字段的声明类型
							// 并根据对象类型设置值
							if (String.class == fieldType) {
								field.set(entity, value);
							} else if (Integer.class == fieldType) {
								field.set(entity, Integer.parseInt(value));
							} else if (Long.class == fieldType) {
								field.set(entity, Long.valueOf(value));
							} else if (Double.class == fieldType) {
								field.set(entity, Double.valueOf(value));
							} else if (Float.class == fieldType) {
								field.set(entity, Float.valueOf(value));
							} else if (Short.class == fieldType) {
								field.set(entity, Short.valueOf(value));
							} else if (Character.class == fieldType) {
								field.set(entity, Character.valueOf(value.charAt(0)));
							}
						} else {
							throw new MyException("第" + (i + 1) + "行第" + (j + 1) + "列不能为空!");// 数据为空提示
						}
					}
					if (entity != null) {
						list.add(entity);
					}
				}
				if (list.size() > 0) {
					System.out.println("Excel导入成功！");
					inputStream.close();
					System.out.println("输入流关闭成功！");
				}
			} else {
				throw new MyException("该表格为空！");// 表格为空提示
			}
		} catch (MyException e) {
			e.printStackTrace();
		} catch (InstantiationException e) {
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return list;
	}

	/**
	 * 将java泛型对象流转成excel文件。
	 *
	 * @param sheetName 工作表的名称
	 * @param sheetSize 每个sheet中数据的行数,此数值必须小于65536
	 * @param output    java输出流
	 * @param list      java泛型对象集合
	 */
	@SuppressWarnings({ "deprecation", "resource" })
	public boolean ExcelExport(List<T> list, String sheetName, int sheetSize, OutputStream output) {
		Field[] allFields = clazz.getDeclaredFields();// 得到所有定义字段
		List<Field> fields = new ArrayList<Field>();
		// 得到所有field并存放到一个list中.
		for (Field field : allFields) {
			if (field.isAnnotationPresent(ExcelIO.class)) {
				fields.add(field);
			}
		}
		HSSFWorkbook workbook = new HSSFWorkbook();// 产生工作薄对象
		// excel2003中每个sheet中最多有65536行
		if (sheetSize > 65536 || sheetSize < 1) {
			sheetSize = 65536;
		}
		double sheetNo = Math.ceil(list.size() / sheetSize);// 取出一共有多少个sheet.
		for (int index = 0; index <= sheetNo; index++) {
			HSSFSheet sheet = workbook.createSheet();// 产生工作表对象
			if (sheetNo == 0) {
				workbook.setSheetName(index, sheetName);
			} else {
				workbook.setSheetName(index, sheetName + index);// 设置工作表的名称.
			}
			HSSFRow row;
			HSSFCell cell;// 产生单元格
			row = sheet.createRow(0);// 产生一行
			// 写入各个字段的列头名称
			for (int i = 0; i < fields.size(); i++) {
				Field field = fields.get(i);
				ExcelIO attr = field.getAnnotation(ExcelIO.class);
				int col = getExcelCol(attr.column());// 获得列号
				cell = row.createCell(col);// 创建列
				cell.setCellType(HSSFCell.CELL_TYPE_STRING);// 设置列中写入内容为String类型
				cell.setCellValue(attr.name());// 写入列名
			}
			int startNo = index * sheetSize;
			int endNo = Math.min(startNo + sheetSize, list.size());
			// 写入各条记录,每条记录对应excel表中的一行
			for (int i = startNo; i < endNo; i++) {
				row = sheet.createRow(i + 1 - startNo);
				T vo = (T) list.get(i); // 得到导出对象.
				for (int j = 0; j < fields.size(); j++) {
					Field field = fields.get(j);// 获得field.
					field.setAccessible(true);// 设置实体类私有属性可访问
					ExcelIO attr = field.getAnnotation(ExcelIO.class);
					try {
						// 根据ExcelIO中设置情况决定是导出全部数据还是只导出标题
						if (attr.isExport()) {
							cell = row.createCell(getExcelCol(attr.column()));// 创建cell
							cell.setCellType(HSSFCell.CELL_TYPE_STRING);
							cell.setCellValue(field.get(vo) == null ? "" : String.valueOf(field.get(vo)));// 如果数据存在就填入,不存在填入空格.
						}
					} catch (IllegalArgumentException e) {
						e.printStackTrace();
					} catch (IllegalAccessException e) {
						e.printStackTrace();
					}
				}
			}
		}
		try {
			output.flush();
			workbook.write(output);
			output.close();
			System.out.println("输出流关闭成功！");
			return true;
		} catch (IOException e) {
			e.printStackTrace();
			return false;
		}
	}

	/**
	 * 得到实体类所有通过注解映射了数据表的字段 递归调用
	 */
	@SuppressWarnings({ "unused", "rawtypes" })
	private List<Field> getMappedFiled(Class clazz, List<Field> fields) {
		if (fields == null) {
			fields = new ArrayList<Field>();
		}
		Field[] allFields = clazz.getDeclaredFields();// 得到所有定义字段
		for (Field field : allFields) {
			if (field.isAnnotationPresent(ExcelIO.class)) {
				fields.add(field);
			}
		}
		if (clazz.getSuperclass() != null && !clazz.getSuperclass().equals(Object.class)) {
			getMappedFiled(clazz.getSuperclass(), fields);
		}
		return fields;
	}

	/**
	 * 将EXCEL中A,B,C,D,E列映射成0,1,2,3
	 */
	private int getExcelCol(String col) {
		col = col.toUpperCase();
		// 从-1开始计算,字母从1开始运算。
		int count = -1;
		char[] cs = col.toCharArray();
		for (int i = 0; i < cs.length; i++) {
			count += (cs[i] - 64) * Math.pow(26, cs.length - 1 - i);
		}
		return count;
	}
}
