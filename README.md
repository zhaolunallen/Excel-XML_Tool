# Excel-XML_Tool
A tool to achieve xml and xlsx file interchange.

# ReadXls
Excel-Xml互转工具： 
该软件在Excel转XML文件的软件基础上加入了XML转Excel的功能，并对原功能进行优化，实现总表模式和分表模式XML与Excel互转的功能。

## 注意事项：  
1.XML文件中，Item和list之间请严格保持list名=item名+“List” ；   
如：   
     ` <ItemExplodeList>  `  
     ` <ItemExplode ID="1001" Rate="100" />  `  
     ` <ItemExplode ID="1002" Rate="100" />  `  
     ` </ItemExplodeList>  `  
2.sheet名由List名定义，因此list不能超过25个字符；  
3.不要使用“True”和“False”字眼，用别的字符替代，“T”和“F”都可以或其他
4.目前软件转出来的xml文件为ANSI格式，如有需要，使用专门用于文本编码转换的软件可转为UTF8，后续更新会把使其直接转为UTF8格式;

