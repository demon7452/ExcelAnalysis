package com.excel.common;

import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
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
			return null;
		}
	} 
    
    /**
	 * 从导入的Excel的输入流中读取文件中的信息
	 * @param inputStream
	 * @param fileName Excel的文件名
	 * @param startRow 从第几行开始
	 * @param columnTotalNum 一共要读取的列数
	 * @param sheetNum 要读取的sheet编号
	 * @return 返回 List<List<String>>
	 * @throws Exception
	 */
	public static List<List<String>> readFromInputStream(InputStream inputStream,String fileName,int startRow,int columnTotalNum,int sheetNum)throws Exception
	{
		List<List<String>> listAll = new ArrayList<>();
		//根据文件名，判断是xls文件还是xlsx文件，返回对应的workbook
		Workbook workbook = createWorkBook(inputStream, fileName);
		if(workbook == null)
			return null;
		
		//获取某一个sheet
		Sheet sheet = workbook.getSheetAt(sheetNum);
		if (sheet == null) 
		{
			return null;
		}
		
		Iterator<Row> rows = sheet.rowIterator();
		while (rows.hasNext())
		{
			List<String> rowList = new ArrayList<>();
			Row row = (Row)rows.next();
			//根据startRow判断跳过几行
			if(row.getRowNum() < startRow-1)
				continue;
			int nullCellNum = 0; //空单元格的数量
			for(int i = 0; i < columnTotalNum; i++)
			{
				Cell cell = row.getCell(i);
				if(null != cell)
				{
					switch (cell.getCellType())
					{
						case Cell.CELL_TYPE_NUMERIC: // 数字,说明：数字类型包含一般类型、日期格式、货币格式、百分比
							rowList.add(handleNumericCell(cell));//处理NUMERIC类型的单元格
                            break;  
                        case Cell.CELL_TYPE_STRING: // 字符串  
                            rowList.add(cell.getStringCellValue()); 
                            break;  
                        case Cell.CELL_TYPE_BOOLEAN: // Boolean  
                        	rowList.add(cell.getBooleanCellValue() + "");
                            break;  
                        case Cell.CELL_TYPE_FORMULA: // 公式
                        	rowList.add(handleFormulaCell(cell));
                            break;  
                        case Cell.CELL_TYPE_BLANK: // 空值
                        	nullCellNum ++;
                        	rowList.add("");
                            break;
                        case Cell.CELL_TYPE_ERROR: // 故障 
                        	nullCellNum ++;
                        	rowList.add("");
                            break;  
                        default:
                        	nullCellNum ++;
                        	rowList.add("");
                            break;
					}
				}
				else
				{
					nullCellNum ++;
					rowList.add("");
				}
			}
			//如果空单元格的数量大于等于所读的单元格，则不加如改行rowList
			if(nullCellNum >= columnTotalNum)
				continue;
			listAll.add(rowList);
		}
		return listAll;
	}
	
	/**
	 * 对公式(Formula)类型的单元格进行相应处理
	 * @param cell
	 * @return
	 */
	public static String handleFormulaCell(Cell cell)
	{
		String result = "";
		try
		{
			double cellValue = cell.getNumericCellValue();
			String style = cell.getCellStyle().getDataFormatString();
			if("#,##".equals(style.substring(0, 4)))
			{
				DecimalFormat format = new DecimalFormat(style.substring(4,style.length()));//设置货币格式，保留两位小数
				result = format.format(cellValue);
			}
			else 
			{
				cell.setCellType(Cell.CELL_TYPE_STRING);
				result = cell.getStringCellValue();
			}
		} catch (Exception e)
		{
//			result = String.valueOf(cell.getRichStringCellValue());
			cell.setCellType(Cell.CELL_TYPE_STRING);
			result = cell.getStringCellValue();
		}
		return result;
	}
	 /**
     * 对NUMERIC类型的单元格进行相应处理
     * 格式化 常规、货币、百分比、日期格式
     * @param cellValue 单元格的NUMERIC值
     * @param styleNum 单元格style
     * @param result 处理后的字符串值
     */
    private static String handleNumericCell(Cell cell)
    {
    	String result = "";
    	try
		{
			double cellValue = cell.getNumericCellValue();
			String style = cell.getCellStyle().getDataFormatString();
			if("#,##".equals(style.substring(0, 4)))
			{
				DecimalFormat format = new DecimalFormat(style.substring(4,style.length()));//设置货币格式，保留两位小数
				result = format.format(cellValue);
			}
			else if("%".equals(style.substring(style.length()-1, style.length())))
			{
//				DecimalFormat format = new DecimalFormat("0.0%");//设置百分比格式，保留小数点后以为
				DecimalFormat format = new DecimalFormat(style.substring(0,style.length()));//设置百分比格式，保留小数点后以为
				result = format.format(cellValue);
			}
			else if("@".equals(style.substring(style.length()-1, style.length())))
			{
				SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");//设置日期格式
//				SimpleDateFormat format = new SimpleDateFormat(style.substring(0,style.length()-2));//设置日期格式
				Date date = DateUtil.getJavaDate(cellValue);//将double类型的日期值转为Date类型
				result = format.format(date);
			}
			else 
			{
				cell.setCellType(Cell.CELL_TYPE_STRING);
				result = cell.getStringCellValue();
			}
		} catch (Exception e)
		{
			cell.setCellType(Cell.CELL_TYPE_STRING);
			result = cell.getStringCellValue();
		}
		return result;
    }
}
