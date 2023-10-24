# PHP实现对学习通Excel课表详细解析以及应用

作者：彭文凤_二C二_ercer
功能：解析超星学习通导出Excel格式的课表
时间：2023年10月22日01:09

站在巨人的肩膀上看世界

## 解析效果

传入参数 tgweek：目标周  day：周几   lesson：第几节课

这里的?day=周二&lesson=1&tgweek=5意思是

第五周->周二->第一节课

![返回的json](https://raw.githubusercontent.com/ccdalao/Analyze_XXT_Schedule_Excel/main/README_Img/3953064178.png)

## Excel课表示例

ok，下面这是我逆天的课表

通过超星学习通导出

![学习通导出Excel课表的教程](https://raw.githubusercontent.com/ccdalao/Analyze_XXT_Schedule_Excel/main/README_Img/772402538.jpg)

![逆天课表](https://raw.githubusercontent.com/ccdalao/Analyze_XXT_Schedule_Excel/main/README_Img/2196839264.png)

## 所需依赖

使用PhpPffice PhpSpreadsheet包 Github：https://github.com/PHPOffice/PhpSpreadsheet

php环境7.2

## 函数列表

获取当课表周数以及当前第几节
`function Get_Now_Current_Weeknumber($startDate, startDate,$targetDate){} function Get_Time_Period($currentTime){}`

通过使用PhpPffice解析Excel
`function Analyze_Excel($File_Address, FileAddress,$sheetname)`

初步解析Excel表源数据
`function Analyze_timetable($tableValue)`

处理Analyze_timetable中的周数信息数据
`function Analyze_TTWeek_data($Data, Data,$NowWeekNuber)`

## 运行流程图
![流程图](https://raw.githubusercontent.com/ccdalao/Analyze_XXT_Schedule_Excel/main/README_Img/20231024090428.png)

## 实现原理

通过对Excel表的解析

利用双重循环，遍历出每天每节课的课程信息

通过对Excel表中每天每节课的课程信息进行解析

获得详细课程信息具体周数

## 拓展使用

可用于消息推送，提示下节课上课，是否有课

可用于自己做出课表小程序，无需点开繁琐学习通

可用于解析多人课表(例如30人)，判断出30人在不同专业班级的情况下，哪一节，大家都没课

## Github❤️

千山万水总是情，给个star行不行

[GitHub：ccdalao 点击跳转](https://github.com/ccdalao/Analyze_XXT_Schedule_Excel)

## 使用示例

![微信通知](https://raw.githubusercontent.com/ccdalao/Analyze_XXT_Schedule_Excel/main/README_Img/660364163.jpg)

![](https://raw.githubusercontent.com/ccdalao/Analyze_XXT_Schedule_Excel/main/README_Img/4104059032.jpg)

通过监控

上课前15分钟自动请求，微信推送下节课的上课地点，课程名称

免得每次上课前担心进错教室，需要从微信退出进入学习通，找到课表，加载半天进行确认教室地点
