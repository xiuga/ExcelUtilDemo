package com.fulan.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author xiang
 * @date 2020年2月16日
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(value = { ElementType.FIELD, ElementType.TYPE })
public @interface ExcelIO {
	/**
	 * 导入指定从第几行开始
	 */
	// public abstract int columnNum() default 0;

	/**
	 * 导出到Excel中的中文列名.
	 */
	public abstract String name();

	/**
	 * 导出每列的顺序,1,2,3,4...对应A,B,C,D...
	 */
	public abstract String column();

	/**
	 * 指定导出全部数据还是只导出标题头
	 */
	public abstract boolean isExport() default true;
}
