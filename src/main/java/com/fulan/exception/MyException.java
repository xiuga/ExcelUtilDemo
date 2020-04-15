package com.fulan.exception;

/**
 * @author xiang
 * @date 2020年2月16日
 */
public class MyException extends Exception{
	private static final long serialVersionUID = 1L;
	public MyException() {
		super();
		// TODO Auto-generated constructor stub
	}

	/**
	 * @param message
	 */
	public MyException(String message) {
		super(message);
		// TODO Auto-generated constructor stub
	}
}
