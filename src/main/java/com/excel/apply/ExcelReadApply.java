package com.excel.apply;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.List;

import com.excel.common.ExcelRead;

public class ExcelReadApply
{
	public static void main(String[] args)
	{
		try
		{
			
			File file = new File("src/main/resources/read.xls");
			String[] aStrings = file.getPath().split("/");
			String fileName = aStrings[aStrings.length-1];
			InputStream inputStream = new FileInputStream(file);
			List<List<String>> allLists = ExcelRead.readFromInputStream(inputStream, fileName, 4, 15, 0);
			for(List<String> list: allLists)
			{
				for(String value : list)
				{
					System.out.print(value + "  ");
				}
				System.out.println("");
			}
			inputStream.close();
		} catch (Exception e)
		{
			// TODO: handle exception
		}

	}
}
