<?php
/**

作者：彭文凤_二C二_ercer
功能：解析超星学习通导出Excel格式的课表
时间：2023年10月22日01:09

站在巨人的肩膀上看世界

**/

/****************引入依赖************************/

//引入依赖 使用PhpPffice PhpSpreadsheet包 Github：https://github.com/PHPOffice/PhpSpreadsheet
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;



/****************基础参数************************/
//基础参数设置
//通过开学日期，给定当前日期，计算当前第几周
//通过作息时间，给定当前时间，计算当前第几节课

$University_Start_Date_Set = "2023-08-28"; //开学第一天的时间戳
$University_Now_Date_Set = "2023-10-16"; //目标日期的时间戳
$University_Now_Time_Set = "14:50"; //目标日期时间的时间戳


/************获取当课表周数以及当前第几节*********/
//Get_Now_Current_Weeknumber通过基础参数，计算当前是第几周，返回格式：int
function Get_Now_Current_Weeknumber($startDate, $targetDate) {
    $startDate = strtotime($startDate);
    $targetDate = strtotime($targetDate);
    $daysDifference = floor(($targetDate - $startDate) / (60 * 60 * 24)); // 计算天数差
    $Now_Current_Weeknumber = floor($daysDifference / 7) + 1; // 计算当前周数（向下取整）并加1
    return $Now_Current_Weeknumber;
}
//Get_Time_Period通过传入给定时间，依次对比数组$periods，计算对比次数，得知当前第几节
function Get_Time_Period($currentTime) {
    //$periods是上课时间表，当前是7节大课的上课下课时间
    $periods = [
        ["08:00",
            "09:20"],
        ["09:40",
            "11:00"],
        ["11:30",
            "12:50"],
        ["13:00",
            "14:20"],
        ["14:40",
            "16:00"],
        ["17:00",
            "18:20"],
        ["18:40",
            "20:00"]
    ];

    foreach ($periods as $index => $period) {
        list($startTime, $endTime) = $period;
        if ($currentTime >= $startTime && $currentTime <= $endTime) {
            return $index + 1; // 返回节次编号，从1开始
        }
    }
    return -1; // 当不在任何节次范围内时返回-1
}


/************通过使用PhpPffice解析Excel***********/
//应当传入需要解析的Excel的文件地址，文件名
//通过两个for循环遍历，获取每天，每节课的课程信息源数据
//【请注意，这里解析的是Excel的源数据，不做任何修改】
function Analyze_Excel($File_Address, $sheetname) {
    // 从指定的文件文件夹加载Excel表格
    $Analyze_Excel = IOFactory::load($File_Address . $sheetname . '.xls');
    $worksheet = $Analyze_Excel->getActiveSheet();

    // 定义每天的课节数
    $daysOfWeek = ["周一",
        "周二",
        "周三",
        "周四",
        "周五",
        "周六",
        "周日"];
    //一天最多有14节小课，但是我们习惯当做把2节小课当做1节课
    $periodsPerDay = 14;

    // 要删除的节次，由于学习通导出Excel课表上 每节课实际上是由2节小课组成，
    // 因此我们只需读取1,3,5,7,9,11,13这7个单元格即可获取每天的七节课
    $periodsToDelete = [2,
        4,
        6,
        8,
        10,
        12,
        14];

    $weeklySchedule = []; // 存储每周的课程表

    // 遍历每天
    for ($dayIndex = 0; $dayIndex < count($daysOfWeek); $dayIndex++) {
        $dayCourses = []; // 存储一天的课程

        // 遍历每节课
        for ($periodIndex = 1; $periodIndex <= $periodsPerDay; $periodIndex++) {
            //这里的判断对应上面需要删除不读取的小节数
            if (!in_array($periodIndex, $periodsToDelete)) {
                $columnIndex = 4 + $dayIndex; // 获取相应的列，从第四列开始
                $cell = $worksheet->getCellByColumnAndRow($columnIndex, $periodIndex + 4); // 获取单元格数据，加4是因为数据从第五行开始
                $courseInfo = $cell->getValue(); // 获取课程信息
                $dayCourses[] = $courseInfo; // 将课程信息存入一天的课程数组
            }
        }

        // 将每天的课程信息存入关联数组，关联数组的键是课程节数
        $dailySchedule = array_combine(range(1, count($dayCourses)), $dayCourses);
        $weeklySchedule[$daysOfWeek[$dayIndex]] = $dailySchedule; // 将每天的课程表存入每周的课程表
    }

    // 将整个每周的课程表转换成JSON格式并返回
    //$jsonOutput = json_encode($weeklySchedule, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
    return $weeklySchedule;
}



/*****************初步解析Excel表源数据**************/
//解析课表内课程详细信息
//$tableValue应当为Analyze_Excel解析出的源数据
//需要给出具体某天某节课的源数据才能解析
/*
数据示例
【源数据】
机械创新设计(理论)\n
长春科技学院\n
尹桂敏【2-8(双)周】\n
2208多媒体教室\n
………………\n
机械制造技术基础(理论)\n
长春科技学院\n
尹桂敏【9-12周】\n
2208多媒体教室
【解析后数据】
array(2) {
  [0]=>
  array(6) {
    ["课程名称"]=>
    string(26) "机械创新设计(理论)"
    ["学校"]=>
    string(18) "长春科技学院"
    ["教师"]=>
    string(26) "尹桂敏【2-8(双)周】"
    ["地点"]=>
    string(19) "2208多媒体教室"
    ["周数"]=>
    string(7) "2,4,6,8"
    ["周类型"]=>
    string(3) "双"
  }
  [1]=>
  array(6) {
    ["课程名称"]=>
    string(32) "机械制造技术基础(理论)"
    ["学校"]=>
    string(18) "长春科技学院"
    ["教师"]=>
    string(22) "尹桂敏【9-12周】"
    ["地点"]=>
    string(19) "2208多媒体教室"
    ["周数"]=>
    string(10) "9,10,11,12"
    ["周类型"]=>
    string(6) "正常"
  }
}

*/
function Analyze_timetable($tableValue) {

    //判断是否有课
    if ($tableValue != NULL) {
        $lines = explode("\n………………\n", $tableValue); // 使用分隔符分割每个课程信息
        $courseInfoArray = [];
        foreach ($lines as $line) {
            $courseInfo = explode("\n", $line);
            // 提取方括号内的周数信息
            preg_match('/【(.+?)】/', $courseInfo[2], $matches);
            $weekInfo = isset($matches[1]) ? $matches[1] : '';

            // 处理周类型
            $weekType = '正常';
            if (preg_match('/(双|单)/', $weekInfo, $typeMatches)) {
                $weekType = $typeMatches[1];
                $weekInfo = preg_replace('/\s*\((双|单)\)\s*/', '', $weekInfo); // 移除周类型信息
            }

            // 解析并处理复杂周数格式
            $weekRanges = [];
            $segments = explode(',', $weekInfo);
            foreach ($segments as $segment) {
                if (strpos($segment, '-') !== false) {
                    list($start, $end) = explode('-', $segment);
                    if ($weekType === '双') {
                        // 当周类型为双时，只选择偶数周
                        for ($i = intval($start); $i <= intval($end); $i++) {
                            if ($i % 2 === 0) {
                                $weekRanges[] = $i;
                            }
                        }
                    } elseif ($weekType === '单') {
                        // 当周类型为单时，只选择奇数周
                        for ($i = intval($start); $i <= intval($end); $i++) {
                            if ($i % 2 !== 0) {
                                $weekRanges[] = $i;
                            }
                        }
                    } else {
                        for ($i = intval($start); $i <= intval($end); $i++) {
                            $weekRanges[] = $i;
                        }
                    }
                } else {
                    $weekRanges[] = intval($segment);
                }
            }
            $formattedWeekInfo = implode(',', $weekRanges);

            $courseInfoArray[] = [
                '课程名称' => $courseInfo[0],
                '学校' => $courseInfo[1],
                '教师' => $courseInfo[2]=preg_replace('/【(.+?)】/', '', $courseInfo[2]),
                '地点' => $courseInfo[3],
                '周数' => $formattedWeekInfo,
                '周类型' => $weekType
            ];
        }
    } else {
        $courseInfoArray[] = [
            '课程名称' => "没课",
            '学校' => "没课",
            '教师' => "没课",
            '地点' => "没课",
            '周数' => "没课",
            '周类型' => "没课"
        ];

    }

    return $courseInfoArray;

}


/*************处理Analyze_timetable中的周数信息数据***********/
//由于学习通Excel课程表上只有9-12周，甚至同一单元格有2门课程在不同的周数中上课
//因此Analyze_timetable中解析的数据无法判断具体某周是否有课，以及上什么课
//需要对周数进行处理，以便获取某周应该上某课 例如8周上机械创新设计(理论)，9周应该上机械制造技术基础(理论)
//需要传入Analyze_timetable的初步解析数据
//需要传入某个周数，例如8，则会解析出第八周的实际课程，如果没有则返回null
function Analyze_TTWeek_data($Data, $NowWeekNuber) {
    
    $foundCourses = array();
    foreach ($Data as $course) {
        $weekNumbers = explode(",", $course["周数"]);
        if (in_array($NowWeekNuber, $weekNumbers)) {
            $foundCourses[] = array(
                "课程名称" => $course["课程名称"],
                "地点" => $course["地点"],
                "老师"=>$course["教师"],
            );
        }
    }

    if (empty($foundCourses)) {
        return null;
    } else {
        return $foundCourses;
    }
}



//解析Excel源文件
$C_Analyze_Excel        =      Analyze_Excel("Excel_ke/", "彭文凤");
//目标周数 
$C_Target_WeekNumber    =      5;
//对源数据初步解析
$C_Analyze_timetable    =      Analyze_timetable($C_Analyze_Excel["周五"]['1']);
//对解析后的数据进行课程周数解析
$C_Analyze_TTWeek_data  =      Analyze_TTWeek_data($C_Analyze_timetable,$C_Target_WeekNumber);


//转为json输出
$C_jsondata = json_encode($C_Analyze_TTWeek_data, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);


echo($C_jsondata);
