**ExcelTool**中的`ExcelUtil`是`poi`封装好的excel导出工具。
- **单表导出**：将一个表格导出；对应方法名`downloadTwoTable2Excel(response, content1, title1, content2, title2, fileName)`
- **双表导出**：将二个表格导出到一个`excel`中，对应的方法名为`downloadTwoTable2Excel(response, content1, title1, content2, title2, fileName);`<br/>
  解释一下参数：<br/>
  `response` 是`HttpServletResponse`类型<br/>
  `content1` `content2`是`List<LinkedHashMap<String,Object>>`类型<br/>
  `title1` `title` 是`excel`的表头<br/>
  `fileName`是保存的文件名<br/>
- **下载个2个表格的效果图：**<br/>
![image](https://github.com/12-09/ExcelTool/blob/master/ExcelTool/WebContent/WEB-INF/images/ExcelTool.png)<br/>
- **多个Sheet页在Excel表中**：将多个数据导入到一个`excel` 中的不同`sheet`中，对应的方法名
`downloadMutilSheetInExcel(HttpServletResponse response,List<List<LinkedHashMap<String, Object>>> content, String[] titles,List<String> sheetNames,String fileName)`
参数`sheetName`是sheet页的名称。
多sheet如图:<br/>
![image](https://github.com/12-09/ExcelTool/blob/master/ExcelTool/WebContent/WEB-INF/images/sheetExcel.png)<br/>

- **调整：**<br/>
    在集合下载的时候，可能只会下载集合中的几个字段，而不会下载所有的字段值,所以对工具类进行了调整，详细见ExcelUtil2.java
                                                                                        
                                                                                        
                                                                                          于2016年9月11日深圳
                                                                                            2016年9月20日更新




