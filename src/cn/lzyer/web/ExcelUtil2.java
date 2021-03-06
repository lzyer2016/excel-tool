package cn.lzyer.web;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
/**
* @ClassName: ExcelUtil2
* @Description: excel表格下载
* 			1.支持1个table下载到一个excel中	
* 			2.支持2个table下载到一个excel中
* 			3.只能下载3万行excel，poi限制是65534和内存限制
* 			4.将数据下载到一个excel表中的多个sheet页中
* @author lzyer 
* @date 2017年2月16日 对工具类增加功能
*
 */
public class ExcelUtil2
{
	
	public static void downloadOneTable2Excel(HttpServletResponse response,
			List<LinkedHashMap<String, Object>> content, String[] title,String[] keys,
			String fileName)
	{
		ByteArrayOutputStream os = new ByteArrayOutputStream();
		try
		{
			createWorkBook(content, title, keys).write(os);
			exportExcel(response, os, fileName);
		} catch(final Exception e)
		{
			e.printStackTrace();
		} 
	}
	
	public static void downloadTwoTable2Excel(HttpServletResponse response,
			List<LinkedHashMap<String, Object>> content1, String[] title1,
			List<LinkedHashMap<String, Object>> content2, String[] title2,
			String[] keys1, String[] keys2,
			String fileName)
	{
		ByteArrayOutputStream os = new ByteArrayOutputStream();
		try
		{
			createWorkBook(content1, title1,content2, title2, keys1, keys2).write(os);
			exportExcel(response, os, fileName);
		} catch(final Exception e)
		{
			e.printStackTrace();
		} 
	}
	/**
	* @Title: downloadMutilSheetInExcel
	* @Description: 下载到一个excel中的多个sheet页中
	* @param response
	* @param content
	* @param titles
	* @param sheetNames
	* @param keys
	* @param fileName	
	* @return void    返回类型
	* @throws
	 */
	public static void downloadMutilSheetInExcel(HttpServletResponse response,
			List<List<LinkedHashMap<String, Object>>> content, String[] titles,
			List<String> sheetNames,String[] keys, String fileName){
		ByteArrayOutputStream os = new ByteArrayOutputStream();
		try
		{
			// 创建excel工作簿
			Workbook workbook = new HSSFWorkbook();
			for(int i=0; i< sheetNames.size(); i++){
				createWorkBook(workbook, content.get(i), titles, keys, sheetNames.get(i));
			}
			workbook.write(os);
			exportExcel(response, os, fileName);
		} catch(final Exception e)
		{
			e.printStackTrace();
		} 
	}
	
	private static void createWorkBook(Workbook workbook,
			List<LinkedHashMap<String, Object>> content, String[] titles,String[] keys,
			String sheetName)
	{
				// 创建第一个sheet（页），并命名
				Sheet sheet = workbook.createSheet(sheetName);
				//定义样式并设置
				CellStyle	cs = workbook.createCellStyle();
				CellStyle	cs2 = workbook.createCellStyle();
				setStyle(workbook, sheet,cs, cs2);
				//生成表格
				generateTable(content, titles, keys, workbook, sheet,cs, cs2, 0);
				setAutoWith(keys,sheet);
	}
	
	private static void exportExcel(HttpServletResponse response, ByteArrayOutputStream os,String fileName)
	{
		BufferedInputStream bis = null;
		BufferedOutputStream bos = null;
		
		try
		{
			byte[] content = os.toByteArray();
			InputStream is = new ByteArrayInputStream(content);
			// 设置response参数，可以打开下载页面
			response.reset();
			response.setContentType("application/vnd.ms-excel;charset=utf-8");
			response.setHeader("Content-Disposition", "attachment;filename="
					+ new String((fileName + ".xls").getBytes(), "iso-8859-1"));
			ServletOutputStream out = response.getOutputStream();

			bis = new BufferedInputStream(is);
			bos = new BufferedOutputStream(out);
			byte[] buff = new byte[2048];
			int bytesRead;
			// Simple read/write loop.
			while (-1 != (bytesRead = bis.read(buff, 0, buff.length)))
			{
				bos.write(buff, 0, bytesRead);
			}
		} catch (Exception e)
		{
			e.printStackTrace();
		}finally{
			 try
			{
				 os.close();
				if (bis != null)
				       bis.close();
				 if (bos != null)
				       bos.close();
			} catch (Exception e)
			{
				e.printStackTrace();
			}
		}
		
	}
	private static Workbook createWorkBook(
			List<LinkedHashMap<String, Object>> content1, String[] title1,
			List<LinkedHashMap<String, Object>> content2, String[] title2,
			String[] keys1, String[] keys2)
	{
				
				// 创建excel工作簿
				Workbook workbook = new HSSFWorkbook();
				// 创建第一个sheet（页），并命名
				Sheet sheet = workbook.createSheet("Sheet1");
				//定义样式并设置
				CellStyle	cs = workbook.createCellStyle();
				CellStyle	cs2 = workbook.createCellStyle();
				setStyle(workbook, sheet, cs, cs2);
				
				//填空第一张表
				generateTable(content1, title1, keys1, workbook, sheet, cs, cs2, 0);
				
				//填充第二张表
				generateTable(content2, title2, keys2, workbook, sheet, cs, cs2, content1.size()+1);
				setAutoWith(keys1,sheet);
				setAutoWith(keys2,sheet);
				return workbook;
	}

	private static void setAutoWith(String[] keys, Sheet sheet)
	{
		for (int i = 0; i < keys.length; i++)
		{
			//自适应宽度
			sheet.autoSizeColumn(i);
		}
	}
	/**
	* @Title: createWorkBook
	* @Description: 创建工作薄
	* @param content
	* @param title
	* @param keys
	* @return	
	* @return Workbook    返回类型
	* @throws
	 */
	private static Workbook createWorkBook(
			List<LinkedHashMap<String, Object>> content, String title[],String[] keys)
	{
		// 创建excel工作簿
		Workbook workbook = new HSSFWorkbook();
		
		// 创建第一个sheet（页），并命名
		Sheet sheet = workbook.createSheet("Sheet1");
		//定义样式并设置
		CellStyle	cs = workbook.createCellStyle();
		CellStyle	cs2 = workbook.createCellStyle();
		setStyle(workbook, sheet,cs, cs2);
		//生成表格
		generateTable(content, title, keys, workbook, sheet,cs, cs2, 0);
		setAutoWith(keys,sheet);
		return workbook;
	}
	/**
	* @Title: setStyle
	* @Description: 设置样式
	* @param wb
	* @param sheet
	* @param cs
	* @param cs2	
	* @return void    返回类型
	* @throws
	 */
	private static void setStyle(Workbook wb, Sheet sheet,CellStyle cs, CellStyle cs2)
	{
		
		// 创建两种字体
		Font f = wb.createFont();
		Font f2 = wb.createFont();

		// 创建第一种字体样式（用于列名）
		f.setFontHeightInPoints((short) 12);
		f.setColor(IndexedColors.BLACK.getIndex());
		f.setBoldweight(Font.BOLDWEIGHT_BOLD);

		// 创建第二种字体样式（用于值）
		f2.setFontHeightInPoints((short) 12);
		f2.setColor(IndexedColors.BLACK.getIndex());

		// Font f3=wb.createFont();
		// f3.setFontHeightInPoints((short) 10);
		// f3.setColor(IndexedColors.RED.getIndex());

		// 设置第一种单元格的样式（用于列名）
		cs.setFont(f);
		cs.setBorderLeft(CellStyle.BORDER_THIN);
		cs.setBorderRight(CellStyle.BORDER_THIN);
		cs.setBorderTop(CellStyle.BORDER_THIN);
		cs.setBorderBottom(CellStyle.BORDER_THIN);
		cs.setAlignment(CellStyle.ALIGN_CENTER);

		// 设置第二种单元格的样式（用于值）
		cs2.setFont(f2);
		cs2.setBorderLeft(CellStyle.BORDER_THIN);
		cs2.setBorderRight(CellStyle.BORDER_THIN);
		cs2.setBorderTop(CellStyle.BORDER_THIN);
		cs2.setBorderBottom(CellStyle.BORDER_THIN);
		cs2.setAlignment(CellStyle.ALIGN_CENTER);
	}
	/**
	* @Title: generateTable
	* @Description: 生成表格
	* @param content
	* @param title
	* @param keys
	* @param wb
	* @param sheet
	* @param cs
	* @param cs2
	* @param start	
	* @return void    返回类型
	* @throws
	 */
	private static void generateTable(List<LinkedHashMap<String, Object>> content,
			String[] title, String[] keys, Workbook wb, Sheet sheet, CellStyle cs, CellStyle cs2,int start)
	{
		// 创建第一行
		Row row = sheet.createRow(start);
		// 设置列名
		for (int i = 0; i < title.length; i++)
		{
			Cell cell = row.createCell(i);
			cell.setCellValue(title[i]);
			cell.setCellStyle(cs);
		}
		// 设置每行每列的值
		for (short offset = 0; offset < content.size(); offset++)
		{
			// 创建一行，在页sheet上
			Row row1 = sheet.createRow(start+offset+1);
			// 在row行上创建一个方格
			for (short j = 0; j < keys.length; j++)
			{
				Cell cell = row1.createCell(j);
				cell.setCellValue(content.get(offset).get(keys[j]) == null ? " " : content
						.get(offset).get(keys[j]).toString());
				cell.setCellStyle(cs2);
			}
		}
	}
	
}


