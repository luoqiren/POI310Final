package com.api.util;

import java.io.File;


public class CommonUtil {
	/**
	 * get current date
	 * @return
	 */
	public static Long getCurrentTime(){
		return System.currentTimeMillis();
	}
	/**
	* <p>Description: </p>
	* @author Qi
	* @date 2017-04-23
	* @param fullPath = path + fileName
	* @return 
	 */
	public static boolean exitsFile(String fullPath){
		return new File(fullPath).exists();
	}
}
