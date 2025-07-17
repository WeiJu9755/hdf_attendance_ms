<?php

//error_reporting(E_ALL); 
//ini_set('display_errors', '1');

session_start();

$memberID = $_SESSION['memberID'];
$powerkey = $_SESSION['powerkey'];

/**
 * PHPExcel
 *
 * Copyright (C) 2006 - 2012 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2012 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.7.8, 2012-10-12
 */

/** Error reporting */
/*
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set("Asia/Taipei");
*/
ini_set('display_errors', FALSE);
ini_set('display_startup_errors', FALSE);
date_default_timezone_set("Asia/Taipei");




//計算請假時間
function calculateLeaveHours($startTime, $endTime)
{
    // 設定工作時間的起點與終點
    $workStartTime = new DateTime('08:00');
    $workEndTime = new DateTime('17:00');

    // 中午休息時間
    $lunchStartTime = new DateTime('12:00');
    $lunchEndTime = new DateTime('13:00');

    // 將請假開始時間與結束時間轉為 DateTime 物件
    $start = new DateTime($startTime);
    $end = new DateTime($endTime);

    // 檢查是否在工作時間範圍內
    if ($start < $workStartTime)
        $start = $workStartTime;
    if ($end > $workEndTime)
        $end = $workEndTime;

    // 計算總請假時數（包含中午時間）
    $interval = $start->diff($end);
    $leaveHours = $interval->h + ($interval->i / 60);

    // 檢查是否跨越中午休息時間，並扣除 1 小時
    if ($start < $lunchEndTime && $end > $lunchStartTime) {
        $leaveHours -= 1; // 扣除 1 小時中午休息時間
    }

    // 將請假時間以半小時為單位計算
    $leaveHours = ceil($leaveHours * 2) / 2;

    return $leaveHours;
}




//載入公用函數
@include_once '/website/include/pub_function.php';


$site_db = "eshop";
$web_id = "sales.eshop";

$team_id = $_GET['team_id'];
$team_id2 = $_GET['team_id2'];
$team_construction_id = $_GET['team_construction_id'];

$start_date = $_GET['start_date'];
$end_date = $_GET['end_date'];

$team_row = getkeyvalue2($site_db . '_info', 'team', "team_id = '$team_id'", 'team_name');
$team_name = $team_row['team_name'];


/*
//檢查是否為管理員及進階會員
$super_admin = "N";
$super_advanced = "N";
$mem_row = getkeyvalue2('memberinfo','member',"member_no = '$memberID'",'admin,advanced,checked,luck,admin_readonly,advanced_readonly');
$super_admin = $mem_row['admin'];
$super_advanced = $mem_row['advanced'];
*/

@include_once("/website/class/" . $site_db . "_info_class.php");


if (PHP_SAPI == 'cli')
    die('This programe should only be run from a Web Browser');

/** Include PHPExcel */
require_once '/website/os/PHPExcel-1.8.1/Classes/PHPExcel.php';
// require_once 'vendor/phpoffice/phpexcel/Classes/PHPExcel.php';
// require_once 'vendor/phpoffice/phpexcel/Classes/PHPExcel/IOFactory.php';


// Create new PHPExcel object
$objPHPExcel = new PHPExcel();
$sheet = $objPHPExcel->getActiveSheet();

// Set document properties
$objPHPExcel->getProperties()->setCreator("PowerSales")
    ->setLastModifiedBy("PowerSales")
    ->setTitle("Office 2007 XLSX Document")
    ->setSubject("Office 2007 XLSX Document")
    ->setDescription("The document for Office 2007 XLSX, generated using PHP classes.")
    ->setKeywords("office 2007 openxml php")
    ->setCategory($team_name . "_" . $start_date . "~" . $end_date . "_團隊支援表");



// 設定起始日期和結束日期
$date1 = new DateTime($start_date);
$date2 = new DateTime($end_date);

// 日期間隔 (以一天為單位)
$day_interval = new DateInterval('P1D');

// 使用 DatePeriod 來生成日期陣列
$period = new DatePeriod($date1, $day_interval, $date2->modify('+1 day'));

// 初始化陣列
$m_header = [];
$m_attendance_day_TOT = [];

// 將日期加入陣列
foreach ($period as $date) {
    $m_header[] = [$date->format('Y-m-d'), $date->format('d')];
    $m_attendance_day_TOT[] = 0;
    $m_attendance_day[] = 0;
}


// 總共天數
$days = Total_days($start_date, $end_date);

// 標題
table_title($team_name,$start_date, $end_date,$days, $sheet);

// 首列表頭欄位
$first_rows =
    [
        'A3' => '序號',
        'B3' => '團隊',
        'C3' => '員工'
    ];
header_first($first_rows, $sheet);



// 列出所有日期
Excel_show_days($start_date, $days, $sheet);


$mDB = "";
$mDB = new MywebDB();

$mDB2 = "";
$mDB2 = new MywebDB();


//取得工單資料
if (!empty($team_construction_id)) {
    if (!empty($team_id2)) {
        $Qry = "SELECT a.dispatch_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day,b.employee_id,c.employee_name,a.team_id,d.team_name 
		,b.team_id as team_id2,e.team_name as team_name2,b.team_construction_id,f.construction_site
		FROM dispatch a
		LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
		LEFT JOIN employee c ON c.employee_id = b.employee_id
		LEFT JOIN team d ON d.team_id = a.team_id
		LEFT JOIN team e ON e.team_id = b.team_id
		LEFT JOIN construction f ON f.construction_id = b.team_construction_id
		WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' 
		AND a.team_id = '$team_id' AND b.team_id is not null AND b.team_id <> '' 
		AND b.team_id = '$team_id2' AND b.team_construction_id = '$team_construction_id'
		GROUP BY b.employee_id
		ORDER BY c.team_id,b.employee_id";
    } else {
        $Qry = "SELECT a.dispatch_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day,b.employee_id,c.employee_name,a.team_id,d.team_name 
		,b.team_id as team_id2,e.team_name as team_name2,b.team_construction_id,f.construction_site
		FROM dispatch a
		LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
		LEFT JOIN employee c ON c.employee_id = b.employee_id
		LEFT JOIN team d ON d.team_id = a.team_id
		LEFT JOIN team e ON e.team_id = b.team_id
		LEFT JOIN construction f ON f.construction_id = b.team_construction_id
		WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' 
		AND a.team_id = '$team_id' AND b.team_id is not null AND b.team_id <> '' 
		AND b.team_construction_id = '$team_construction_id'
		GROUP BY b.employee_id
		ORDER BY c.team_id,b.employee_id";
    }
} else {
    if (!empty($team_id2)) {
        $Qry = "SELECT a.dispatch_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day,b.employee_id,c.employee_name,a.team_id,d.team_name 
		,b.team_id as team_id2,e.team_name as team_name2,b.team_construction_id,f.construction_site
		FROM dispatch a
		LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
		LEFT JOIN employee c ON c.employee_id = b.employee_id
		LEFT JOIN team d ON d.team_id = a.team_id
		LEFT JOIN team e ON e.team_id = b.team_id
		LEFT JOIN construction f ON f.construction_id = b.team_construction_id
		WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' 
		AND a.team_id = '$team_id' AND b.team_id is not null AND b.team_id <> '' 
		AND b.team_id = '$team_id2'
		GROUP BY b.employee_id
		ORDER BY c.team_id,b.employee_id";
    } else {
        $Qry = "SELECT a.dispatch_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day,b.employee_id,c.employee_name,a.team_id,d.team_name 
		,b.team_id as team_id2,e.team_name as team_name2,b.team_construction_id,f.construction_site
		FROM dispatch a
		LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
		LEFT JOIN employee c ON c.employee_id = b.employee_id
		LEFT JOIN team d ON d.team_id = a.team_id
		LEFT JOIN team e ON e.team_id = b.team_id
		LEFT JOIN construction f ON f.construction_id = b.team_construction_id
		WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' 
		AND a.team_id = '$team_id' AND b.team_id is not null AND b.team_id <> '' 
		GROUP BY b.employee_id
		ORDER BY c.team_id,b.employee_id";
    }
}


$mDB->query($Qry);

$total = $mDB->rowCount();
if ($total > 0) {


    $members = [];
    $member_teamwork = [];
    $seq = 0;

    while ($row = $mDB->fetchRow(2)) {
        // 將員工資料加入 $members 陣列
        $members[] = [
            'team' => $row['team_name'],
            'name' => $row['employee_name'],
            'employee_id' => $row['employee_id'],
        ];

        if (!empty($row['employee_id'])) {
            $seq++;

            // 取得該員工的詳細資料
            $Qry2 = "SELECT 
                    a.dispatch_id, 
                    a.dispatch_date, 
                    YEAR(a.dispatch_date) AS dispatch_year, 
                    MONTH(a.dispatch_date) AS dispatch_month, 
                    DAY(a.dispatch_date) AS dispatch_day, 
                    b.employee_id, 
                    b.manpower, 
                    b.attendance_day, 
                    b.attendance_status, 
                    b.transition, 
                    b.transition_start, 
                    b.transition_end, 
                    f.construction_site, 
                    c.team_name, 
                    d.team_name AS transition_team_name
                FROM dispatch a
                LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
                LEFT JOIN team c ON c.team_id = b.team_id
                LEFT JOIN team d ON d.team_id = b.transition_team_id
        		LEFT JOIN construction f ON f.construction_id = b.team_construction_id
                WHERE 
                    a.dispatch_date BETWEEN '$start_date' AND '$end_date' 
                    AND a.ConfirmSending = 'Y' 
                    AND b.employee_id = '{$row['employee_id']}'
                ORDER BY b.employee_id";

            $mDB2->query($Qry2);
            while ($row2 = $mDB2->fetchRow(2)) {

                if (!empty($row2['construction_site'])) {
                    $construction_site_str = $row2['team_name']."\n".$row2['construction_site'];
                } else {
                    $construction_site_str = $row2['team_name'];
                }

                $member_teamwork[] = [
                    'id' => $row2['employee_id'],
                    'dispatch_day' => $row2['dispatch_day'],
                    'attendance_status' => $row2['attendance_status'],
                    'attendance_time' => $row2['attendance_day'],
                    'construction_site' => $construction_site_str,
                    'transition' => $row2['transition'],
                    'transition_team' => $row2['transition_team_name'],
                    'transition_start' => $row2['transition_start'],
                    'transition_end' => $row2['transition_end'],
                ];
            }
        }
    }
}
// 員工資訊
member_info($members, $sheet);
// 支援報表
support_report($start_date, $end_date, $member_teamwork, $sheet);

// 最後列表頭欄位
header_last($days, $sheet);

$mDB2->remove();
$mDB->remove();

// 計算日期範圍的天數
function Total_days($start_date, $end_date)
{
    $start_timestamp = strtotime($start_date);
    $end_timestamp = strtotime($end_date);
    $days_difference = ceil(($end_timestamp - $start_timestamp) / (60 * 60 * 24)) + 1; // 加 1 包含開始日
    return $days_difference;
}

// 標題
function table_title($name,$start_date, $end_date,$days, $sheet)
{
    $columnLetter = \PHPExcel_Cell::stringFromColumnIndex(3 + $days);
    $sheet->mergeCells('A1:' . $columnLetter . '2');
    $sheet->setCellValue('A1',$name . $start_date."~". $end_date . '_團隊支援表');

    // 設定儲存格高度為40
    $sheet->getRowDimension(1)->setRowHeight(20);
    $sheet->getRowDimension(2)->setRowHeight(20);

    // 設定儲存格文字對齊
    $sheet->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle('A1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

    // 設定邊框
    $styleArray = [
        'borders' => [
            'allborders' => [
                'style' => PHPExcel_Style_Border::BORDER_THIN, // 粗線
                'color' => ['argb' => 'FF000000'], // 黑色
            ],
        ],
    ];
    //設置字型大小
    $sheet->getStyle('A1')->getFont()->setSize(30)->setBold(true);
    $sheet->getStyle('A1:' . $columnLetter . '2')->applyFromArray($styleArray);
}
// 首列表頭欄位
function header_first($first_rows, $sheet)
{
    foreach ($first_rows as $key => $value) {

        if ($value == 'A3') {
            $sheet->getColumnDimension('A')->setWidth(3);
        } else {
            $sheet->getColumnDimension('B')->setWidth(15);
            $sheet->getColumnDimension('C')->setWidth(15);
        }
        // 設定儲存格值
        $sheet->setCellValue($key, $value);

        // 設定文字置中
        $sheet->getStyle($key)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle($key)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

        // 設定粗體字
        $sheet->getStyle($key)->getFont()->setBold(true);

        // 設定底色
        $sheet->getStyle($key)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
        $sheet->getStyle($key)->getFill()->getStartColor()->setRGB('7FDBFF'); // 底色為 #7FDBFF

        // 設定邊框
        $styleArray = [
            'borders' => [
                'allborders' => [
                    'style' => PHPExcel_Style_Border::BORDER_THIN, // 細線
                    'color' => ['argb' => 'FF000000'], // 黑色
                ],
            ],
        ];
        $sheet->getStyle($key)->applyFromArray($styleArray);
    }
}

// 最後列表頭欄位"合計"
function header_last($days, $sheet)
{
    $columnLetter = \PHPExcel_Cell::stringFromColumnIndex(3 + $days);

    // 設定文字置中
    $sheet->getStyle($columnLetter . '3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle($columnLetter . '3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

    // 設定底色
    $sheet->getStyle($columnLetter . '3')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
    $sheet->getStyle($columnLetter . '3')->getFill()->getStartColor()->setRGB('7FDBFF'); // 底色為 #7FDBFF

    // 設定粗體字
    $sheet->getStyle($columnLetter . '3')->getFont()->setBold(true);

    // 設定邊框（全部細線）
    $styleBorders = [
        'borders' => [
            'allborders' => [
                'style' => PHPExcel_Style_Border::BORDER_THIN, // 細線
                'color' => ['argb' => 'FF000000'], // 黑色
            ],
        ],
    ];
    $sheet->getStyle($columnLetter . '3')->applyFromArray($styleBorders);

    // 設定內容為 "合計"
    $sheet->setCellValue($columnLetter . '3', '合計');

    // 設定列寬度為 15
    $sheet->getColumnDimension($columnLetter)->setWidth(15);
}


// 在工作表中填入日期
function Excel_show_days($start_date, $days, $sheet)
{
    $date = new DateTime($start_date);

    if ($sheet === null) {
        throw new Exception("Invalid sheet object.");
    }

    for ($i = 0; $i < $days; $i++) {
        $current_date = (clone $date)->modify("+$i days")->format('d'); // 每次加一天
        $columnLetter = PHPExcel_Cell::stringFromColumnIndex(3 + $i); // 計算欄位字母

        // 設定日期
        $sheet->setCellValue($columnLetter . '3', $current_date);

        // 設定文字置中
        $sheet->getStyle($columnLetter . '3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle($columnLetter . '3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

        // 設定底色
        $sheet->getStyle($columnLetter . '3')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
        $sheet->getStyle($columnLetter . '3')->getFill()->getStartColor()->setRGB('7FDBFF'); // 設置底色

        // 設定邊框（細線）
        $styleBorders = [
            'borders' => [
                'allborders' => [
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => ['argb' => 'FF000000'], // 黑色
                ],
            ],
        ];
        $sheet->getStyle($columnLetter . '3')->applyFromArray($styleBorders);
    }
}



//員工資訊 
function member_info($members, $sheet, $startRow = 4)
{
    // 計算成員數量
    $memberCount = count($members);

    for ($i = 0; $i < $memberCount; $i++) {
        $currentRow = $startRow + ($i * 3); // 每個成員佔用 3 行

        // 合併儲存格 (A 列，用於顯示 ID)
        $sheet->mergeCells("A{$currentRow}:A" . ($currentRow + 2));
        $sheet->setCellValue("A{$currentRow}", $i + 1);

        // 設置團隊 (B 列)
        $sheet->mergeCells("B{$currentRow}:B" . ($currentRow + 2));
        $sheet->setCellValue("B{$currentRow}", $members[$i]['team']);

        // 設置成員名稱 (C 列)
        $sheet->mergeCells("C{$currentRow}:C" . ($currentRow + 2));
        $sheet->setCellValue("C{$currentRow}", $members[$i]['name']);

        // 水平與垂直居中
        $sheet->getStyle("A{$currentRow}:C" . ($currentRow + 2))
            ->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle("A{$currentRow}:C" . ($currentRow + 2))
            ->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

        // 設置字體大小為 12
        $sheet->getStyle("A{$currentRow}:C" . ($currentRow + 2))
            ->getFont()->setSize(12);

        // 設置四邊框為黑色細線
        $sheet->getStyle("A{$currentRow}:C" . ($currentRow + 2))
            ->applyFromArray([
                'alignment' => [
                    'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                    'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
                ],
                'borders' => [
                    'allborders' => [
                        'style' => PHPExcel_Style_Border::BORDER_THIN,
                        'color' => ['argb' => 'FF000000'],
                    ],
                ],
                'font' => ['bold' => true],
            ]);
    }

    // 合計列
    $totalRow = $startRow + ($memberCount * 3); // 計算合計所在的行
    $sheet->mergeCells("A{$totalRow}:C{$totalRow}"); // 合併 A~C 列
    $sheet->setCellValue("A{$totalRow}", "合計"); // 設置值為「合計」

    // 設置居中對齊
    $sheet->getStyle("A{$totalRow}:C{$totalRow}")
        ->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle("A{$totalRow}:C{$totalRow}")
        ->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

    // 設置加粗字體
    $sheet->getStyle("A{$totalRow}:C{$totalRow}")->getFont()->setBold(true);

    // 設置字體大小為 12
    $sheet->getStyle("A{$totalRow}:C{$totalRow}")
        ->getFont()->setSize(12);

    // 設置四邊框為黑色細線
    $sheet->getStyle("A{$totalRow}:C{$totalRow}")->applyFromArray([
        'alignment' => [
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'allborders' => [
                'style' => PHPExcel_Style_Border::BORDER_THICK,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
        'font' => ['bold' => true],
    ]);

    // 設置合計列底色為 FFDC00
    $sheet->getStyle("A{$totalRow}:C{$totalRow}")
        ->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
    $sheet->getStyle("A{$totalRow}:C{$totalRow}")
        ->getFill()->getStartColor()->setRGB('FFDC00');
}
// 支援團隊報表
function support_report($start_date, $end_date, $member_teamwork, $sheet)
{
    if ($sheet === null) {
        throw new Exception("Invalid sheet object.");
    }
    $date = new DateTime($start_date);
    $start_day = (int) $date->format('d');

    // 計算總天數
    $total_days = Total_days($start_date, $end_date);

    // 整理資料：按 ID 與 dispatch_day 分組
    $teamwork_map = [];
    foreach ($member_teamwork as $entry) {
        $teamwork_map[$entry['id']][$entry['dispatch_day']] = $entry;
    }

    // 開始列偏移（假設日期從第 4 列開始）
    $column_offset = 3;
    $initial_row = 4; // 每個 ID 起始行
    $block_height = 3; // 每個 ID 的區塊高度

    // 計算最後一欄的字母
    $last_column_index = $column_offset + $total_days;
    $last_column_letter = \PHPExcel_Cell::stringFromColumnIndex($last_column_index);

    // 初始化每日加總數組
    $daily_totals = array_fill(1, $total_days, 0);
    foreach ($teamwork_map as $id => $dispatches) {
        $current_row = $initial_row;
        $total_hours = 0;

        foreach ($dispatches as $dispatch_day => $entry) {
            $columnIndex = $dispatch_day - $start_day; // 對應到日期的列
            $columnLetter = \PHPExcel_Cell::stringFromColumnIndex($column_offset + $columnIndex);

            // 設置 Attendance Time
            if (isset($entry['attendance_time']) && $entry['attendance_status'] === "支援") {
                $sheet->setCellValue("{$columnLetter}{$current_row}", round($entry['attendance_time']/8,4));
                $total_hours += round($entry['attendance_time']/8,4); // 累加總時數
                $daily_totals[$dispatch_day - $start_day + 1] += round($entry['attendance_time']/8,4); // 每日加總
            }

            // 設置 Attendance Status 與 Construction Site
            if (isset($entry['attendance_status']) && $entry['attendance_status'] === "支援" && isset($entry['construction_site'])) {
                $sheet->setCellValue("{$columnLetter}" . ($current_row + 1), $entry['attendance_status'] . ":" . $entry['construction_site']);

                // 啟用換行功能
                $sheet->getStyle("{$columnLetter}" . ($current_row + 1))->getAlignment()->setWrapText(true);
            }

            // 設置 Transition 與 Transition Team
            //if (isset($entry['transition']) && $entry['transition'] === "Y" && isset($entry['transition_team'])) {
            if (isset($entry['transition']) && $entry['transition'] === "Y") {
                $transition_hours = round(calculateLeaveHours($entry['transition_start'], $entry['transition_end'])/8,4);
                $sheet->setCellValue("{$columnLetter}" . ($current_row + 2), "轉場:" . $entry['transition_team'] . $transition_hours);
                $daily_totals[$dispatch_day - $start_day + 1] += $transition_hours; // 每日加總
            }
        }

        // 在最後一欄輸入總時數並合併儲存格（垂直兩格）
        $total_hours_cell = "{$last_column_letter}{$current_row}";
        $sheet->setCellValue($total_hours_cell, (string) ($total_hours));
        $sheet->mergeCells("{$total_hours_cell}:{$last_column_letter}" . ($current_row + 2)); // 合併下兩格

        // 每個 ID 的資料結束後，移動到下一區塊
        $initial_row += $block_height;
    }

    // 設置最後一列的每日加總
    $total_row = $initial_row; // 設定最後一列的位置
    $total_all = 0;
    for ($day = $start_day; $day < $start_day + $total_days; $day++) {
        $columnLetter = \PHPExcel_Cell::stringFromColumnIndex($column_offset + ($day - $start_day));
        $sheet->setCellValue("{$columnLetter}{$total_row}", (string) ($daily_totals[$day - $start_day + 1]));
        // 設置合計列底色為 FFDC00
        $sheet->getStyle("{$columnLetter}{$total_row}")
            ->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
        $sheet->getStyle("{$columnLetter}{$total_row}")
            ->getFill()->getStartColor()->setRGB('FFDC00');
        $total_all += $daily_totals[$day - $start_day + 1];
    }
    $columnLetter = \PHPExcel_Cell::stringFromColumnIndex($column_offset + ($total_days));
    $sheet->setCellValue("{$columnLetter}{$total_row}", $total_all);
    // 設置合計列底色為 FFDC00
    $sheet->getStyle("{$columnLetter}{$total_row}")
        ->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
    $sheet->getStyle("{$columnLetter}{$total_row}")
        ->getFill()->getStartColor()->setRGB('FFDC00');

 

    // 設置儲存格樣式（判斷有無值）
    for ($row = 4; $row <= $total_row; $row++) {
        for ($col = $column_offset; $col <= $last_column_index; $col++) {
            $columnLetter = \PHPExcel_Cell::stringFromColumnIndex($col);

            // 嘗試讀取儲存格值（包括公式計算結果）
            $cellValue = $sheet->getCell("{$columnLetter}{$row}")->getCalculatedValue();

            if (!empty($cellValue)) {
                // 儲存格有值，設置寬度為 15
                $sheet->getColumnDimension($columnLetter)->setWidth(15);
            } else {
                // 儲存格無值，設置寬度為 3
                $sheet->getColumnDimension($columnLetter)->setWidth(3);
            }

            $sheet->getRowDimension($row)->setRowHeight(20);

            // 設置樣式
            $styleArray = [
                'alignment' => [
                    'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                    'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
                ],
                'borders' => [
                    'allborders' => [
                        'style' => $row === $total_row
                            ? PHPExcel_Style_Border::BORDER_THICK
                            : PHPExcel_Style_Border::BORDER_THIN,
                        'color' => ['argb' => 'FF000000'],
                    ],
                ],
                'font' => ['bold' => $row === $total_row],
            ];

            $sheet->getStyle("{$columnLetter}{$row}")->applyFromArray($styleArray);
        }
    }
}

// Rename worksheet
$objPHPExcel->getActiveSheet()->setTitle("團隊支援表");


$xlsx_filename = $team_name . "_" . $start_date . "~" . $end_date . "_團隊支援表.xls";


// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename=' . $xlsx_filename);
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header('Pragma: public'); // HTTP/1.0

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
exit;