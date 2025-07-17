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



function number_format2($num,$dec) {
	if ($num <> 0)
		if ($num > 0.0001) {
			$retval = number_format($num,$dec);
		} else {
			$retval = "";
		}
	else 
		$retval = "";
		
	return $retval;
}

//計算請假時間
function calculateLeaveHours($startTime, $endTime) {
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
    if ($start < $workStartTime) $start = $workStartTime;
    if ($end > $workEndTime) $end = $workEndTime;

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

$company_id = $_GET['company_id'];
$m_team_id = $_GET['team_id'];

$m_team_name = "";

$start_date = $_GET['start_date'];
$end_date = $_GET['end_date'];

$company_row = getkeyvalue2($site_db.'_info','company',"company_id = '$company_id'",'company_name');
$company_name = $company_row['company_name'];

/*
//檢查是否為管理員及進階會員
$super_admin = "N";
$super_advanced = "N";
$mem_row = getkeyvalue2('memberinfo','member',"member_no = '$memberID'",'admin,advanced,checked,luck,admin_readonly,advanced_readonly');
$super_admin = $mem_row['admin'];
$super_advanced = $mem_row['advanced'];
*/

@include_once("/website/class/".$site_db."_info_class.php");


if (PHP_SAPI == 'cli')
	die('This programe should only be run from a Web Browser');

/** Include PHPExcel */
require_once '/website/os/PHPExcel-1.8.1/Classes/PHPExcel.php';


// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

// Set document properties
if (!empty($m_team_id)) {

	$team_row = getkeyvalue2($site_db.'_info','team',"team_id = '$m_team_id'",'team_name');
	$m_team_name = $team_row['team_name'];
	
	$objPHPExcel->getProperties()->setCreator("PowerSales")
								->setLastModifiedBy("PowerSales")
								->setTitle("Office 2007 XLSX Document")
								->setSubject("Office 2007 XLSX Document")
								->setDescription("The document for Office 2007 XLSX, generated using PHP classes.")
								->setKeywords("office 2007 openxml php")
								->setCategory($company_name."_".$m_team_name."_".$start_date."~".$end_date."_出勤表");
} else {
	$objPHPExcel->getProperties()->setCreator("PowerSales")
								->setLastModifiedBy("PowerSales")
								->setTitle("Office 2007 XLSX Document")
								->setSubject("Office 2007 XLSX Document")
								->setDescription("The document for Office 2007 XLSX, generated using PHP classes.")
								->setKeywords("office 2007 openxml php")
								->setCategory($company_name."_".$start_date."~".$end_date."_出勤表");
}


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
    $m_header[] = [$date->format('Y-m-d'),$date->format('d')];
    $m_attendance_day_TOT[] = 0;
	$m_attendance_day[] = 0;
}




$mDB = "";
$mDB = new MywebDB();

$mDB2 = "";
$mDB2 = new MywebDB();

//取得工單資料
if (!empty($team_id)) {
	$Qry="SELECT a.dispatch_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day,b.employee_id,c.employee_name,c.team_id,d.team_name FROM dispatch a
	LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
	LEFT JOIN employee c ON c.employee_id = b.employee_id
	LEFT JOIN team d ON d.team_id = c.team_id
	WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' AND a.company_id = '$company_id' AND a.team_id = '$team_id'
	GROUP BY b.employee_id
	ORDER BY c.team_id,b.employee_id";
} else {
	$Qry="SELECT a.dispatch_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day,b.employee_id,c.employee_name,c.team_id,d.team_name FROM dispatch a
	LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
	LEFT JOIN employee c ON c.employee_id = b.employee_id
	LEFT JOIN team d ON d.team_id = c.team_id
	WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' AND a.company_id = '$company_id'
	GROUP BY b.employee_id
	ORDER BY c.team_id,b.employee_id";
}

$mDB->query($Qry);

$total = $mDB->rowCount();

$line = 1;

if ($total > 0) {



	//設置對齊
	$objPHPExcel->getActiveSheet()->getStyle('1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

	//匯出主要資料表
	$objPHPExcel->setActiveSheetIndex(0)
				->setCellValue('A1', '序號')
				->setCellValue('B1', '團隊')
				->setCellValue('C1', '員工')
				;

	//設置行列高度
	$objPHPExcel->getActiveSheet()->getRowDimension(1)->setRowHeight(25);

	//設置邊框線及顏色			
	$objPHPExcel->getActiveSheet()->getStyle('A1')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	$objPHPExcel->getActiveSheet()->getStyle('A1')->getBorders()->getAllBorders()->getColor()->setRGB('000000');
	$objPHPExcel->getActiveSheet()->getStyle('B1')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	$objPHPExcel->getActiveSheet()->getStyle('B1')->getBorders()->getAllBorders()->getColor()->setRGB('000000');
	$objPHPExcel->getActiveSheet()->getStyle('C1')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	$objPHPExcel->getActiveSheet()->getStyle('C1')->getBorders()->getAllBorders()->getColor()->setRGB('000000');
	//設置垂及及水平對齊
	$objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objPHPExcel->getActiveSheet()->getStyle('B1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objPHPExcel->getActiveSheet()->getStyle('B1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objPHPExcel->getActiveSheet()->getStyle('C1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objPHPExcel->getActiveSheet()->getStyle('C1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	//設置寬度
	$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
	$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(20);
	$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);
	//設置字型大小
	$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setSize(12)->setBold(true);
	$objPHPExcel->getActiveSheet()->getStyle('B1')->getFont()->setSize(12)->setBold(true);
	$objPHPExcel->getActiveSheet()->getStyle('C1')->getFont()->setSize(12)->setBold(true);
	//設置底色
	$objPHPExcel->getActiveSheet()->getStyle('A1:C1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
	$objPHPExcel->getActiveSheet()->getStyle('A1:C1')->getFill()->getStartColor()->setRGB('7FDBFF');


	$i_count = count($m_header);

	$startColumn = 'C';
	$startColumnIndex = \PHPExcel_Cell::columnIndexFromString($startColumn); // 起始欄位轉為數字
	$endColumnIndex = $startColumnIndex + $i_count;
	$row = 1; // 要套用樣式的列

	$i = 0;
	for ($colIndex = $startColumnIndex; $colIndex < $endColumnIndex; $colIndex++) {
		$column = \PHPExcel_Cell::stringFromColumnIndex($colIndex); // 轉回欄位字母
		$cell = $column . $row;
		$objPHPExcel->setActiveSheetIndex(0)->setCellValue($cell, $m_header[$i][1]);
		//設置邊框線及顏色			
		$objPHPExcel->getActiveSheet()->getStyle($cell)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		$objPHPExcel->getActiveSheet()->getStyle($cell)->getBorders()->getAllBorders()->getColor()->setRGB('000000');
		//設置垂及及水平對齊
		$objPHPExcel->getActiveSheet()->getStyle($cell)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->getStyle($cell)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
		//設置寬度
		$objPHPExcel->getActiveSheet()->getColumnDimension($column)->setWidth(12);
		//設置字型大小
		$objPHPExcel->getActiveSheet()->getStyle($cell)->getFont()->setSize(12)->setBold(true);
		//設置底色
		$objPHPExcel->getActiveSheet()->getStyle($cell)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
		$objPHPExcel->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setRGB('7FDBFF');
		$i++;
	}
	$column = \PHPExcel_Cell::stringFromColumnIndex($colIndex); // 轉回欄位字母
	$cell = $column . $row;
	$objPHPExcel->setActiveSheetIndex(0)->setCellValue($cell, "合計");
	//設置邊框線及顏色			
	$objPHPExcel->getActiveSheet()->getStyle($cell)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	$objPHPExcel->getActiveSheet()->getStyle($cell)->getBorders()->getAllBorders()->getColor()->setRGB('000000');
	//設置垂及及水平對齊
	$objPHPExcel->getActiveSheet()->getStyle($cell)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objPHPExcel->getActiveSheet()->getStyle($cell)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	//設置寬度
	$objPHPExcel->getActiveSheet()->getColumnDimension($column)->setWidth(20);
	//設置字型大小
	$objPHPExcel->getActiveSheet()->getStyle($cell)->getFont()->setSize(12)->setBold(true);
	//設置底色
	$objPHPExcel->getActiveSheet()->getStyle($cell)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
	$objPHPExcel->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setRGB('7FDBFF');


	$seq = 0;
	while ($row=$mDB->fetchRow(2)) {
		$dispatch_id = $row['dispatch_id'];
		$dispatch_year = $row['dispatch_year'];
		$dispatch_month = $row['dispatch_month'];
		$dispatch_day = $row['dispatch_day'];
		$employee_id = $row['employee_id'];
		$employee_name = $row['employee_name'];
		$team_id = $row['team_id'];
		$team_name = $row['team_name'];

		if (!empty($employee_id)) {
			$seq++;
		
			/*
			//再取得各員工的資料
			$Qry2="SELECT a.dispatch_id,a.dispatch_date,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day
				,b.employee_id,b.manpower,b.attendance_day,b.attendance_status,b.leave,b.leave_start,b.leave_end,b.work_overtime,b.work_overtime_start,b.work_overtime_end
				,b.attendance_end_status,b.attendance_end,b.transition,b.transition_team_id,b.transition_construction_id,b.transition_start,b.transition_end
				FROM dispatch a
				LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
				WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' AND b.employee_id = '$employee_id'
				ORDER BY b.employee_id";
			*/

			//再取得各員工的資料
			$Qry2="SELECT a.dispatch_id,a.dispatch_date,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day
				,b.employee_id,b.manpower,b.attendance_day,b.attendance_status,b.leave,b.leave_start,b.leave_end,b.work_overtime,b.work_overtime_start,b.work_overtime_end
				,b.attendance_end_status,b.attendance_end,b.transition,b.transition_team_id,b.transition_construction_id,b.transition_start,b.transition_end
				FROM dispatch a
				LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
				WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' AND b.employee_id = '$employee_id'
				ORDER BY b.employee_id";


			$mDB2->query($Qry2);
			$SUMMARY = 0;
			if ($mDB2->rowCount() > 0) {

				$m_attendance_day = [];
				foreach ($period as $date) {
					$m_attendance_day[] = 0;
				}

				$m_leave = [];
				foreach ($period as $date) {
					$m_leave[] = '';
				}


				while ($row2=$mDB2->fetchRow(2)) {

					$DAY = $row2['dispatch_date'];

					$i_count = count($m_attendance_day);
					for ($i = 0; $i < $i_count; $i++) {

						if ($DAY == $m_header[$i][0]) {

							$status_str = "";

							if ($row2['attendance_status'] == "颱風假") {
								$status_str .= $row2['attendance_status'];
							}

							if ($row2['transition'] == "Y") {
								//計算轉場時數/日
								$transition_start = $row2['transition_start'];
								$transition_end = $row2['transition_end'];
	
								$transition_hours = round(calculateLeaveHours($transition_start, $transition_end)/8,4);
	
								$attendance_end = substr($row2['attendance_end'],0,5);
								$status_str .= "轉場".$transition_hours;
							}
	

							if (!empty($row2['leave'])) {
								//計算請假時數/日
								$leave_start = $row2['leave_start'];
								$leave_end = $row2['leave_end'];

								$leave_hours = round(calculateLeaveHours($leave_start, $leave_end)/8,4);

								$status_str .= $row2['leave'].$leave_hours;
							}

							if ($row2['attendance_end_status'] == "提早收工") {
								$attendance_end = substr($row2['attendance_end'],0,5);
								$status_str .= $attendance_end.$row2['attendance_end_status'];
							}
	
							if (!empty($row2['work_overtime'])) {
								//計算加班時數/日
								// 定義兩個時間
								$work_overtime_start = new DateTime($DAY." ".$row2['work_overtime_start']);
								$work_overtime_end = new DateTime($DAY." ".$row2['work_overtime_end']);

								//比較時間大小
								if ($work_overtime_end < $work_overtime_start) {
									//如果 $work_overtime_end < $work_overtime_start 則 work_overtime_end 要加一天
									$work_overtime_end = $work_overtime_end->modify('+1 day');
								}

								// 計算時間差
								$interval = $work_overtime_start->diff($work_overtime_end);

								// 將時間差轉換為總分鐘數
								$total_minutes = ($interval->days * 24 * 60) + ($interval->h * 60) + $interval->i;

								// 將分鐘轉換為小時，並四捨五入到小數點1位
								$work_overtime_hours = round($total_minutes / 60 /8, 4);

								$status_str .= $row2['work_overtime'].$work_overtime_hours;
							}


							$m_leave[$i] = $status_str;



							$m_attendance_day[$i] = round($row2['attendance_day']/8,4);
							$SUMMARY = $SUMMARY+$m_attendance_day[$i];	//橫向加總
							//累計加總
							$m_attendance_day_TOT[$i] = $m_attendance_day_TOT[$i]+$m_attendance_day[$i];
							
						}

						if ($m_attendance_day[$i] == 0) {
							$m_attendance_day[$i] = number_format2($m_attendance_day[$i],4);
						}

					}

				}
			}
		}

		$attendance_day_TOTAL = $attendance_day_TOTAL+$SUMMARY;

		$i_count = count($m_attendance_day_TOT);
		for ($i = 0; $i < $i_count; $i++) {
			if ($m_attendance_day_TOT[$i] == 0) {
				$m_attendance_day_TOT[$i] = number_format2($m_attendance_day_TOT[$i],4);
			}
		}

		$fmt_SUMMARY = number_format2($SUMMARY,4);

		$fmt_attendance_day_TOTAL = number_format2($attendance_day_TOTAL,4);
		


		
		$line++;
		$row = $line;
		$row2 = $line+1;

		$cell = 'A' . $line;
		$mergeRange = "A{$row}:A{$row2}";
		$objPHPExcel->setActiveSheetIndex(0)->mergeCells($mergeRange);
		$objPHPExcel->setActiveSheetIndex(0)->setCellValue($cell, $seq);
		//設置邊框線及顏色			
		$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getBorders()->getAllBorders()->getColor()->setRGB('000000');
		//設置垂及及水平對齊
		$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$cell = 'B' . $line;
		$mergeRange = "B{$row}:B{$row2}";
		$objPHPExcel->setActiveSheetIndex(0)->mergeCells($mergeRange);
		$objPHPExcel->setActiveSheetIndex(0)->setCellValue($cell, $team_name." ".$team_id);
		//設置邊框線及顏色			
		$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getBorders()->getAllBorders()->getColor()->setRGB('000000');
		//設置垂及及水平對齊
		$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		$cell = 'C' . $line;
		$mergeRange = "C{$row}:C{$row2}";
		$objPHPExcel->setActiveSheetIndex(0)->mergeCells($mergeRange);
		$objPHPExcel->setActiveSheetIndex(0)->setCellValue($cell, $employee_name." ".$employee_id);
		//設置邊框線及顏色			
		$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getBorders()->getAllBorders()->getColor()->setRGB('000000');
		//設置垂及及水平對齊
		$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
		
		$i_count = count($m_attendance_day);

		$startColumn = 'C';
		$startColumnIndex = \PHPExcel_Cell::columnIndexFromString($startColumn); // 起始欄位轉為數字
		$endColumnIndex = $startColumnIndex + $i_count;
	
		$i = 0;
		for ($colIndex = $startColumnIndex; $colIndex < $endColumnIndex; $colIndex++) {
			$column = \PHPExcel_Cell::stringFromColumnIndex($colIndex); // 轉回欄位字母
			$cell = $column . $row;
			$objPHPExcel->setActiveSheetIndex(0)->setCellValue($cell, $m_attendance_day[$i]);
			//設置邊框線及顏色			
			$objPHPExcel->getActiveSheet()->getStyle($cell)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objPHPExcel->getActiveSheet()->getStyle($cell)->getBorders()->getAllBorders()->getColor()->setRGB('000000');
			//設置垂及及水平對齊
			$objPHPExcel->getActiveSheet()->getStyle($cell)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objPHPExcel->getActiveSheet()->getStyle($cell)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
			$i++;
		}
		
		$i = 0;
		for ($colIndex = $startColumnIndex; $colIndex < $endColumnIndex; $colIndex++) {
			$column = \PHPExcel_Cell::stringFromColumnIndex($colIndex); // 轉回欄位字母
			$cell2 = $column . $row2;
			$objPHPExcel->setActiveSheetIndex(0)->setCellValue($cell2, $m_leave[$i]);
			//設置邊框線及顏色			
			$objPHPExcel->getActiveSheet()->getStyle($cell2)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
			$objPHPExcel->getActiveSheet()->getStyle($cell2)->getBorders()->getAllBorders()->getColor()->setRGB('000000');
			//設置垂及及水平對齊
			$objPHPExcel->getActiveSheet()->getStyle($cell2)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objPHPExcel->getActiveSheet()->getStyle($cell2)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
			$i++;
		}
		
		$column = \PHPExcel_Cell::stringFromColumnIndex($colIndex); // 轉回欄位字母
		$cell = $column . $row;
		$mergeRange = "{$column}{$row}:{$column}{$row2}";
		$objPHPExcel->setActiveSheetIndex(0)->mergeCells($mergeRange);
		$objPHPExcel->setActiveSheetIndex(0)->setCellValue($cell, $fmt_SUMMARY);
		//設置邊框線及顏色			
		$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getBorders()->getAllBorders()->getColor()->setRGB('000000');
		//設置垂及及水平對齊
		$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
		

		$line++;
	

	}

	//顯示全部總和
	$row = $line+1;

	//設置行列高度
	$objPHPExcel->getActiveSheet()->getRowDimension($row)->setRowHeight(25);

	$cell = 'A' . $row;
	$mergeRange = "A{$row}:C{$row}";
	$objPHPExcel->setActiveSheetIndex(0)->mergeCells($mergeRange);
	$objPHPExcel->setActiveSheetIndex(0)->setCellValue($cell, "合計");
	//設置邊框線及顏色			
	$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getBorders()->getAllBorders()->getColor()->setRGB('000000');
	//設置垂及及水平對齊
	$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	//設置字型大小
	$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getFont()->setSize(12)->setBold(true);
	//設置底色
	$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
	$objPHPExcel->getActiveSheet()->getStyle($mergeRange)->getFill()->getStartColor()->setRGB('FFDC00');


	$i_count = count($m_attendance_day_TOT);

	$startColumn = 'C';
	$startColumnIndex = \PHPExcel_Cell::columnIndexFromString($startColumn); // 起始欄位轉為數字
	$endColumnIndex = $startColumnIndex + $i_count;

	$i = 0;
	for ($colIndex = $startColumnIndex; $colIndex < $endColumnIndex; $colIndex++) {
		$column = \PHPExcel_Cell::stringFromColumnIndex($colIndex); // 轉回欄位字母
		$cell = $column . $row;
		$objPHPExcel->setActiveSheetIndex(0)->setCellValue($cell, $m_attendance_day_TOT[$i]);
		//設置邊框線及顏色			
		$objPHPExcel->getActiveSheet()->getStyle($cell)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		$objPHPExcel->getActiveSheet()->getStyle($cell)->getBorders()->getAllBorders()->getColor()->setRGB('000000');
		//設置垂及及水平對齊
		$objPHPExcel->getActiveSheet()->getStyle($cell)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$objPHPExcel->getActiveSheet()->getStyle($cell)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
		//設置字型大小
		$objPHPExcel->getActiveSheet()->getStyle($cell)->getFont()->setSize(12)->setBold(true);
		//設置底色
		$objPHPExcel->getActiveSheet()->getStyle($cell)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
		$objPHPExcel->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setRGB('FFDC00');
		$i++;
	}

	$column = \PHPExcel_Cell::stringFromColumnIndex($colIndex); // 轉回欄位字母
	$cell = $column . $row;
	$objPHPExcel->setActiveSheetIndex(0)->setCellValue($cell, $fmt_attendance_day_TOTAL);
	//設置邊框線及顏色			
	$objPHPExcel->getActiveSheet()->getStyle($cell)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	$objPHPExcel->getActiveSheet()->getStyle($cell)->getBorders()->getAllBorders()->getColor()->setRGB('000000');
	//設置垂及及水平對齊
	$objPHPExcel->getActiveSheet()->getStyle($cell)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objPHPExcel->getActiveSheet()->getStyle($cell)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	//設置字型大小
	$objPHPExcel->getActiveSheet()->getStyle($cell)->getFont()->setSize(12)->setBold(true);
	//設置底色
	$objPHPExcel->getActiveSheet()->getStyle($cell)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
	$objPHPExcel->getActiveSheet()->getStyle($cell)->getFill()->getStartColor()->setRGB('FFDC00');



}


$mDB2->remove();
$mDB->remove();


// Rename worksheet
$objPHPExcel->getActiveSheet()->setTitle("出勤表");


if (!empty($m_team_id)) {
	$xlsx_filename = $company_name."_".$m_team_name."_".$start_date."~".$end_date."_出勤表.xls";
} else {
	$xlsx_filename = $company_name."_".$start_date."~".$end_date."_出勤表.xls";
}


// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename='.$xlsx_filename);
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
exit;




?>
