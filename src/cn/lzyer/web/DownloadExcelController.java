package cn.lzyer.web;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

@Controller
@RequestMapping("/download")
public class DownloadExcelController
{
	
	@RequestMapping(value="/hello")
	public String hello()
	{
		return "hello";
	}
	
	@RequestMapping("/excel")
	@ResponseBody
	public void downloadExcel(HttpServletRequest request, HttpServletResponse response)
	{
		//模拟数据库中的数据
		String[] title1 = {"序号","姓名","日期"};
		List<LinkedHashMap<String,Object>> content1 = new ArrayList<LinkedHashMap<String,Object>>();
		
		for(int i=0; i<5; i++)
		{
			LinkedHashMap<String, Object> map = new LinkedHashMap<String, Object>();
			map.put("number", i+"");
			map.put("name", "用户"+i);
			map.put("date", new SimpleDateFormat("yyyy-MM-dd").format(new Date()));
			content1.add(map);
		}
		
		String[] title2 = {"序号","姓名","日期","课程","班级","工资","学校","省份","城市","兴趣","年龄","地址","职位",
				"银行类型","卡类型","账户","余额","开户地","创建时间","更新时间","修改人","备注"};
		List<LinkedHashMap<String,Object>> content2 = new ArrayList<LinkedHashMap<String,Object>>();
		
		for(int i=0; i<3; i++)
		{
			LinkedHashMap<String, Object> map = new LinkedHashMap<String, Object>();
			map.put("number", i+"");
			map.put("name", "用户"+i);
			map.put("date", new SimpleDateFormat("yyyy-MM-dd").format(new Date()));
			map.put("course", "计算机"+i);
			map.put("class", "计科"+i);
			map.put("money", "2000"+i);
			map.put("school", "大学"+i);
			map.put("province", "广东省"+i);
			map.put("city", "城市"+i);
			map.put("interest", "coding"+i);
			map.put("age", "2"+i);
			map.put("address", "中国广东省深圳市南山区白石洲沙河镇上白石xxxxx"+i);
			map.put("position", "程序员"+i);
			map.put("banktype", "招商银行"+i);
			map.put("cardtype", "储蓄卡"+i);
			map.put("bankaccount", "42802"+i);
			map.put("yue", "12345"+i);
			
			map.put("openadd", "白石洲支行"+i);
			map.put("createtime", new SimpleDateFormat("yyyy-MM-dd").format(new Date()));
			map.put("updatetime", null);
			map.put("updatepeople", null);
			map.put("remark", null);
			content2.add(map);
		}
		
		//下载文件
		//ExcelUtil.downloadTwoTable2Excel(response, content1, title1, content2, title2, "学校人员信息");
		ExcelUtil.downloadOneTable2Excel(response, content1, title1, "用户信息");
	}
	/**
	* @Title: downloadMutilSheetExcel
	* @Description: 测试下载到多个sheet页中
	* @param @param request
	* @param @param response    参数
	* @return void    返回类型
	* @throws
	 */
	@RequestMapping("/sheetexcel")
	@ResponseBody
	public void downloadMutilSheetExcel(HttpServletRequest request, HttpServletResponse response){
		//模拟数据库中的数据
				String[] title1 = {"序号","姓名","日期"};
				List<List<LinkedHashMap<String,Object>>> contents = new ArrayList<List<LinkedHashMap<String,Object>>>();
				for(int j=0; j<5; j++){
					List<LinkedHashMap<String,Object>> content1 = new ArrayList<LinkedHashMap<String,Object>>();
					for(int i=0; i<5; i++)
					{
						LinkedHashMap<String, Object> map = new LinkedHashMap<String, Object>();
						map.put("number", i+"");
						map.put("name", "用户"+i);
						map.put("date", new SimpleDateFormat("yyyy-MM-dd").format(new Date()));
						content1.add(map);
					}
					contents.add(content1);
				}
				List<String> sheetNames = new ArrayList<String>();
				sheetNames.add("七月用户数据");
				sheetNames.add("八月用户数据");
				sheetNames.add("九月用户数据");
				sheetNames.add("十月用户数据");
				sheetNames.add("十一月用户数据");
				ExcelUtil.downloadMutilSheetInExcel(response, contents, title1, sheetNames, "用户数据");
	}
	@RequestMapping("/exportV2")
	public void downloadExcelV2(HttpServletResponse response){
				//模拟数据库中的数据
				//String[] title1 = {"序号","姓名","日期"};
				String[] title1 = {"序号","姓名","城市"};
				List<LinkedHashMap<String,Object>> content1 = new ArrayList<LinkedHashMap<String,Object>>();
				
				for(int i=0; i<5; i++)
				{
					LinkedHashMap<String, Object> map = new LinkedHashMap<String, Object>();
					map.put("number", i+"");
					map.put("name", "用户"+i);
					map.put("date", new SimpleDateFormat("yyyy-MM-dd").format(new Date()));
					content1.add(map);
				}
				
				String[] title2 = {"序号","姓名","日期","课程","班级","工资","学校","省份","城市","兴趣","年龄","地址","职位",
						"银行类型","卡类型","账户","余额","开户地","创建时间","更新时间","修改人","备注"};
				List<LinkedHashMap<String,Object>> content2 = new ArrayList<LinkedHashMap<String,Object>>();
				
				for(int i=0; i<3; i++)
				{
					LinkedHashMap<String, Object> map = new LinkedHashMap<String, Object>();
					map.put("number", i+"");
					map.put("name", "用户"+i);
					map.put("date", new SimpleDateFormat("yyyy-MM-dd").format(new Date()));
					map.put("course", "计算机"+i);
					map.put("class", "计科"+i);
					map.put("money", "2000"+i);
					map.put("school", "大学"+i);
					map.put("province", "广东省"+i);
					map.put("city", "城市"+i);
					map.put("interest", "coding"+i);
					map.put("age", "2"+i);
					map.put("address", "中国广东省深圳市南山区白石洲沙河镇上白石xxxxx"+i);
					map.put("position", "程序员"+i);
					map.put("banktype", "招商银行"+i);
					map.put("cardtype", "储蓄卡"+i);
					map.put("bankaccount", "42802"+i);
					map.put("yue", "12345"+i);
					
					map.put("openadd", "白石洲支行"+i);
					map.put("createtime", new SimpleDateFormat("yyyy-MM-dd").format(new Date()));
					map.put("updatetime", null);
					map.put("updatepeople", null);
					map.put("remark", null);
					content2.add(map);
				}
				
				//下载文件
				//ExcelUtil.downloadTwoTable2Excel(response, content1, title1, content2, title2, "学校人员信息");
				String keys[] = {"number","name","createtime"};
				ExcelUtil2.downloadOneTable2Excel(response, content2, title1,keys, "用户信息");
	}
}


