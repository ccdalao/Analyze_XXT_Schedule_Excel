# PHP实现对学习通Excel课表详细解析以及应用

作者：彭文凤_二C二_ercer
功能：解析超星学习通导出Excel格式的课表
时间：2023年10月22日01:09

站在巨人的肩膀上看世界

## 解析效果

传入参数 tgweek：目标周  day：周几   lesson：第几节课

这里的?day=周二&lesson=1&tgweek=5意思是

第五周->周二->第一节课

![返回的json](http://blog.ercer.cn/usr/uploads/2023/10/3953064178.png)

## Excel课表示例

ok，下面这是我逆天的课表

通过超星学习通导出

![学习通导出Excel课表的教程](http://blog.ercer.cn/usr/uploads/2023/10/772402538.jpg)

![逆天课表](http://blog.ercer.cn/usr/uploads/2023/10/2196839264.png)

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

<svg zoomAndPan="magnify" width="830px" viewBox="0 0 830 1229" version="1.1" style="width: 80%; height: 80%; background: rgb(255, 255, 255);" preserveAspectRatio="none" height="1229px" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns="http://www.w3.org/2000/svg"><defs></defs><g><rect y="15" x="290" width="247" style="stroke:none;stroke-width:1.0;" id="_title" height="26.2969" fill="none"></rect><text y="32.9951" x="295" textLength="237" lengthAdjust="spacing" font-weight="bold" font-size="14" font-family="sans-serif" fill="#000000">解析超星学习通导出Excel格式的课表</text><ellipse style="stroke:#222222;stroke-width:1.0;" ry="10" rx="10" fill="#222222" cy="85.9951" cx="115"></ellipse><rect y="115.9951" x="21" width="188" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="137.1338" x="31" textLength="168" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">传入 tgweek, day, 和 lesson</text><rect y="169.9639" x="51" width="128" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="191.1025" x="61" textLength="108" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">调用Analyze_Excel</text><rect y="1087.4326" x="69" width="92" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="1108.5713" x="79" textLength="72" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">返回课程信息</text><rect y="1141.4014" x="66.5" width="97" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="1162.54" x="76.5" textLength="77" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">输出JSON数据</text><ellipse style="stroke:#222222;stroke-width:1.0;" ry="11" rx="11" fill="none" cy="1206.3701" cx="115"></ellipse><ellipse style="stroke:#111111;stroke-width:1.0;" ry="6" rx="6" fill="#222222" cy="1206.3701" cx="115"></ellipse><line y2="1217.3701" y1="50.042" x2="15" x1="15" style="stroke:#000000;stroke-width:1.5;"></line><rect y="223.9326" x="263" width="100" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="245.0713" x="273" textLength="80" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">加载Excel表格</text><rect y="277.9014" x="255" width="116" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="299.04" x="265" textLength="96" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">遍历每天和每节课</text><rect y="331.8701" x="267" width="92" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="353.0088" x="277" textLength="72" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">跳过对应小课</text><rect y="385.8389" x="255" width="116" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="406.9775" x="265" textLength="96" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">获取相应的列和行</text><rect y="439.8076" x="261" width="104" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="460.9463" x="271" textLength="84" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">获取单元格数据</text><rect y="493.7764" x="267" width="92" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="514.915" x="277" textLength="72" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">获取课程信息</text><rect y="547.7451" x="267" width="92" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="568.8838" x="277" textLength="72" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">存储课程信息</text><rect y="601.7139" x="219" width="188" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="622.8525" x="229" textLength="168" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">将每天的课程信息存入关联数组</text><rect y="655.6826" x="249" width="128" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="676.8213" x="259" textLength="108" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">将每周的课程表存储</text><rect y="709.6514" x="235.5" width="155" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="730.79" x="245.5" textLength="135" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">调用Analyze_timetable</text><line y2="1217.3701" y1="50.042" x2="213" x1="213" style="stroke:#000000;stroke-width:1.5;"></line><rect y="763.6201" x="460.5" width="92" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="784.7588" x="470.5" textLength="72" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">判断是否有课</text><rect y="817.5889" x="448.5" width="116" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="838.7275" x="458.5" textLength="96" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">分割每个课程信息</text><rect y="871.5576" x="430.5" width="152" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="892.6963" x="440.5" textLength="132" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">解析并处理复杂周数格式</text><rect y="925.5264" x="448.5" width="116" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="946.665" x="458.5" textLength="96" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">存储解析后的数据</text><rect y="979.4951" x="417" width="179" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="1000.6338" x="427" textLength="159" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">调用Analyze_TTWeek_data</text><line y2="1217.3701" y1="50.042" x2="411" x1="411" style="stroke:#000000;stroke-width:1.5;"></line><rect y="1033.4639" x="630" width="152" style="stroke:#181818;stroke-width:0.5;" ry="12.5" rx="12.5" height="33.9688" fill="#F1F1F1"></rect><text y="1054.6025" x="640" textLength="132" lengthAdjust="spacing" font-size="12" font-family="sans-serif" fill="#000000">解析指定周数的实际课程</text><line y2="1217.3701" y1="50.042" x2="600" x1="600" style="stroke:#000000;stroke-width:1.5;"></line><line y2="1217.3701" y1="50.042" x2="810" x1="810" style="stroke:#000000;stroke-width:1.5;"></line><line y2="115.9951" y1="95.9951" x2="115" x1="115" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="111,105.9951,115,115.9951,119,105.9951,115,109.9951" fill="#181818"></polygon><line y2="169.9639" y1="149.9639" x2="115" x1="115" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="111,159.9639,115,169.9639,119,159.9639,115,163.9639" fill="#181818"></polygon><line y2="1141.4014" y1="1121.4014" x2="115" x1="115" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="111,1131.4014,115,1141.4014,119,1131.4014,115,1135.4014" fill="#181818"></polygon><line y2="1195.3701" y1="1175.3701" x2="115" x1="115" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="111,1185.3701,115,1195.3701,119,1185.3701,115,1189.3701" fill="#181818"></polygon><line y2="277.9014" y1="257.9014" x2="313" x1="313" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="309,267.9014,313,277.9014,317,267.9014,313,271.9014" fill="#181818"></polygon><line y2="331.8701" y1="311.8701" x2="313" x1="313" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="309,321.8701,313,331.8701,317,321.8701,313,325.8701" fill="#181818"></polygon><line y2="385.8389" y1="365.8389" x2="313" x1="313" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="309,375.8389,313,385.8389,317,375.8389,313,379.8389" fill="#181818"></polygon><line y2="439.8076" y1="419.8076" x2="313" x1="313" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="309,429.8076,313,439.8076,317,429.8076,313,433.8076" fill="#181818"></polygon><line y2="493.7764" y1="473.7764" x2="313" x1="313" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="309,483.7764,313,493.7764,317,483.7764,313,487.7764" fill="#181818"></polygon><line y2="547.7451" y1="527.7451" x2="313" x1="313" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="309,537.7451,313,547.7451,317,537.7451,313,541.7451" fill="#181818"></polygon><line y2="601.7139" y1="581.7139" x2="313" x1="313" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="309,591.7139,313,601.7139,317,591.7139,313,595.7139" fill="#181818"></polygon><line y2="655.6826" y1="635.6826" x2="313" x1="313" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="309,645.6826,313,655.6826,317,645.6826,313,649.6826" fill="#181818"></polygon><line y2="709.6514" y1="689.6514" x2="313" x1="313" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="309,699.6514,313,709.6514,317,699.6514,313,703.6514" fill="#181818"></polygon><line y2="817.5889" y1="797.5889" x2="506.5" x1="506.5" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="502.5,807.5889,506.5,817.5889,510.5,807.5889,506.5,811.5889" fill="#181818"></polygon><line y2="871.5576" y1="851.5576" x2="506.5" x1="506.5" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="502.5,861.5576,506.5,871.5576,510.5,861.5576,506.5,865.5576" fill="#181818"></polygon><line y2="925.5264" y1="905.5264" x2="506.5" x1="506.5" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="502.5,915.5264,506.5,925.5264,510.5,915.5264,506.5,919.5264" fill="#181818"></polygon><line y2="979.4951" y1="959.4951" x2="506.5" x1="506.5" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="502.5,969.4951,506.5,979.4951,510.5,969.4951,506.5,973.4951" fill="#181818"></polygon><line y2="208.9326" y1="203.9326" x2="115" x1="115" style="stroke:#181818;stroke-width:1.0;"></line><line y2="208.9326" y1="208.9326" x2="313" x1="115" style="stroke:#181818;stroke-width:1.0;"></line><line y2="223.9326" y1="208.9326" x2="313" x1="313" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="309,213.9326,313,223.9326,317,213.9326,313,217.9326" fill="#181818"></polygon><line y2="748.6201" y1="743.6201" x2="313" x1="313" style="stroke:#181818;stroke-width:1.0;"></line><line y2="748.6201" y1="748.6201" x2="506.5" x1="313" style="stroke:#181818;stroke-width:1.0;"></line><line y2="763.6201" y1="748.6201" x2="506.5" x1="506.5" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="502.5,753.6201,506.5,763.6201,510.5,753.6201,506.5,757.6201" fill="#181818"></polygon><line y2="1018.4639" y1="1013.4639" x2="506.5" x1="506.5" style="stroke:#181818;stroke-width:1.0;"></line><line y2="1018.4639" y1="1018.4639" x2="706" x1="506.5" style="stroke:#181818;stroke-width:1.0;"></line><line y2="1033.4639" y1="1018.4639" x2="706" x1="706" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="702,1023.4639,706,1033.4639,710,1023.4639,706,1027.4639" fill="#181818"></polygon><line y2="1072.4326" y1="1067.4326" x2="706" x1="706" style="stroke:#181818;stroke-width:1.0;"></line><line y2="1072.4326" y1="1072.4326" x2="115" x1="706" style="stroke:#181818;stroke-width:1.0;"></line><line y2="1087.4326" y1="1072.4326" x2="115" x1="115" style="stroke:#181818;stroke-width:1.0;"></line><polygon style="stroke:#181818;stroke-width:1.0;" points="111,1077.4326,115,1087.4326,119,1077.4326,115,1081.4326" fill="#181818"></polygon><text y="66.75" x="93.5" textLength="41" lengthAdjust="spacing" font-size="18" font-family="sans-serif" fill="#000000">User</text><text y="66.75" x="249.5" textLength="125" lengthAdjust="spacing" font-size="18" font-family="sans-serif" fill="#000000">Analyze_Excel</text><text y="66.75" x="424.5" textLength="162" lengthAdjust="spacing" font-size="18" font-family="sans-serif" fill="#000000">Analyze_timetable</text><text y="66.75" x="605" textLength="200" lengthAdjust="spacing" font-size="18" font-family="sans-serif" fill="#000000">Analyze_TTWeek_data</text></g></svg>

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

![微信通知](http://blog.ercer.cn/usr/uploads/2023/10/660364163.jpg)

![](http://blog.ercer.cn/usr/uploads/2023/10/4104059032.jpg)

通过监控

上课前15分钟自动请求，微信推送下节课的上课地点，课程名称

免得每次上课前担心进错教室，需要从微信退出进入学习通，找到课表，加载半天进行确认教室地点
