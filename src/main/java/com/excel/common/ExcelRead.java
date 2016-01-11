package com.excel.common;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * common methods to read excel
 * 从Excel读取数据
 * @author xiong
 *
 */
public class ExcelRead
{
	/*  * 判断是xls文件还是xlsx文件  */
	/**
	 * 根据读取的文件名的后缀名判断是 xls 或者 xlsx 文件
	 * 返回相应的
	 * @param inputStream
	 * @param fileName 读取的文件名
	 * @return
	 * @throws IOException
	 */
    public static Workbook createWorkBook(InputStream inputStream,String fileName) throws Exception
    {
    	if(fileName.toLowerCase().endsWith("xls"))
    	{
    		return new HSSFWorkbook(inputStream);
    	}
    	else if(fileName.toLowerCase().endsWith("xlsx"))
    	{
    		return new XSSFWorkbook(inputStream);
		}
    	else {
			throw new Exception();
		}

//		if (fileName.toLowerCase().endsWith("xls")) {
//			try {
//				return new HSSFWorkbook(inputStream);
//			} catch (Exception e) {
//				return new XSSFWorkbook(inputStream);
//			}
//		}
//		if (fileName.toLowerCase().endsWith("xlsx")) {
//			try {
//				return new XSSFWorkbook(inputStream);
//			} catch (Exception e) {
//				return new HSSFWorkbook(inputStream);
//			}
//		}
//		return null;
	} 
}
