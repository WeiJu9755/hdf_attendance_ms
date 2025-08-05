<?php

//error_reporting(E_ALL); 
//ini_set('display_errors', '1');

session_start();

$memberID = $_SESSION['memberID'];
$powerkey = $_SESSION['powerkey'];


require_once '/website/os/Mobile-Detect-2.8.34/Mobile_Detect.php';
$detect = new Mobile_Detect;

if (!($detect->isMobile() && !$detect->isTablet())) {
	$isMobile = "0";
} else {
	$isMobile = "1";
}


$m_location = "/website/smarty/templates/" . $site_db . "/" . $templates;
$m_pub_modal = "/website/smarty/templates/" . $site_db . "/pub_modal";


function number_format2($num, $dec)
{
	if ($num <> 0)
		if ($num > 0.0001) {
			$retval = number_format($num, $dec);
		} else {
			$retval = "";
		} else
		$retval = "";

	return $retval;
}

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

@include_once("/website/class/" . $site_db . "_info_class.php");



$fm = $_GET['fm'];
$ch = $_GET['ch'];
//$project_id = $_GET['project_id'];
//$auth_id = $_GET['auth_id'];

$company_id = $_GET['company_id'];

$select_company_DB = "";
$select_company_DB = new MywebDB();

$Qry_select_company = "SELECT company_id,company_name FROM `company` WHERE company_id ='$company_id'";
$select_company_DB->query($Qry_select_company);
$company_row = $select_company_DB->fetchRow(2);

$company_name = $company_row['company_name'];

$start_date = $_GET['start_date'];
$end_date = $_GET['end_date'];

if (!isset($_GET['start_date']))
	$start_date = date('Y-m-d');

if (!isset($_GET['end_date']))
	$end_date = date('Y-m-d');


//檢查是否為管理員及進階會員
$super_admin = "N";
$super_advanced = "N";
$mem_row = getkeyvalue2('memberinfo', 'member', "member_no = '$memberID'", 'admin,advanced');
$super_admin = $mem_row['admin'];
$super_advanced = $mem_row['advanced'];





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
	->setCategory($company_name . "_" . $start_date . "~" . $end_date . "工地狀況查詢");




// 設定起始日期和結束日期
$date1 = new DateTime($start_date);
$date2 = new DateTime($end_date);

// 日期間隔 (以一天為單位)
$day_interval = new DateInterval('P1D');

// 使用 DatePeriod 來生成日期陣列
$period = new DatePeriod($date1, $day_interval, $date2->modify('+1 day'));

// 初始化陣列
$m_header = [];
$m_manpower_total = [];
$m_attendance_day_total = [];

// 將日期加入陣列
foreach ($period as $date) {
	$m_header[] = [$date->format('Y-m-d'), $date->format('d')];
	$m_manpower_total[] = 0;
	$m_attendance_day_total[] = 0;
	$m_manpower[] = 0;
	$m_attendance_day[] = 0;
}

// 總共天數
$days = Total_days($start_date, $end_date);

// 標題
table_title($company_name, $start_date, $end_date, $days, $sheet);

// 首列表頭欄位
$first_rows =
	[
		'A3' => '序號',
		'B3' => '工地',
	];
header_first($first_rows, $sheet);


// 列出所有日期
Excel_show_days($start_date, $days, $sheet);

// 最後列表頭欄位
header_last($days, $sheet);

// 輸出日期陣列
//print_r($dateArray);
//exit;

// 初始化資料庫連線物件
$mDB = "";
$mDB = new MywebDB();

$mDB2 = "";
$mDB2 = new MywebDB();
$get_construction_id = trim($_GET['construction_id'] ?? '');
// 查詢所有需要處理的工地資料
if($_GET['company_id']!=""){
$Qry="SELECT a.dispatch_id,a.company_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day
,b.construction_id,c.construction_site,b.building,b.household,b.floor,b.attendance_status,b.team_id,d.team_name,b.manpower,b.workinghours FROM dispatch a
LEFT JOIN dispatch_construction b ON b.dispatch_id = a.dispatch_id
LEFT JOIN construction c ON c.construction_id = b.construction_id
LEFT JOIN team d ON d.team_id = b.team_id
WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' AND a.company_id = '$company_id'";



$Qry .="GROUP BY b.construction_id
		ORDER BY b.construction_id";

}else{
	$Qry="SELECT a.dispatch_id,a.company_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day
,b.construction_id,c.construction_site,b.building,b.household,b.floor,b.attendance_status,b.team_id,d.team_name,b.manpower,b.workinghours FROM dispatch a
LEFT JOIN dispatch_construction b ON b.dispatch_id = a.dispatch_id
LEFT JOIN construction c ON c.construction_id = b.construction_id
LEFT JOIN team d ON d.team_id = b.team_id
WHERE b.construction_id = '{$_GET['construction_id']}'";



$Qry .="GROUP BY b.construction_id
		ORDER BY b.construction_id";
}

// 執行查詢並儲存結果
$mDB->query($Qry);
$construction_sites = [];

// 將查詢結果存入工地陣列
while ($row = $mDB->fetchRow(2)) {
	$construction_sites[] = [
		'construction_id' => $row['construction_id'],
		'construction_site' => $row['construction_site'],
        'get_company_id' => $row['company_id'],
	];
}

// 初始化報表行偏移量
$row_offset = 0;
// 用於儲存每日總計統計
$daily_totals_summary = [];
$selected_status = trim($_GET['attendance_status'] ?? '');
// 逐一處理每個工地
foreach ($construction_sites as $site) {
    $construction_id = $site['construction_id'];
    $site_name = $site['construction_site'];
    $get_company_id = $site['get_company_id'];

    // 查詢單一工地的詳細派工資料
    $Qry2 = "SELECT 
        a.dispatch_id, a.dispatch_date,
        b.construction_id, c.construction_site,
        b.attendance_status, b.team_id, d.team_name,
        b.manpower, b.workinghours
    FROM dispatch a
    LEFT JOIN dispatch_construction b ON b.dispatch_id = a.dispatch_id
    LEFT JOIN construction c ON c.construction_id = b.construction_id
    LEFT JOIN team d ON d.team_id = b.team_id
    WHERE a.dispatch_date >= '$start_date' 
      AND a.dispatch_date <= '$end_date' 
      AND a.ConfirmSending = 'Y' 
      AND b.construction_id = '$construction_id'
      AND a.company_id = '$get_company_id'";

      if(!empty($selected_status)){
			$Qry2 .= " AND b.attendance_status = '$selected_status'";
		}
    $Qry2 .="ORDER BY a.dispatch_date";

    // 執行查詢
    $mDB2->query($Qry2);
    // 用於儲存單一工地的分組報表資料
    $construction_report_grouped = [];

    // 處理查詢結果，依日期分組
    while ($row2 = $mDB2->fetchRow(2)) {
        $dispatch_date = $row2['dispatch_date'];
        $attendance_status = $row2['attendance_status'];
        $team_name = $row2['team_name'];
        $manpower = (int) $row2['manpower'];
        // 將工時轉換為小數格式（以 8 小時為基準）
        $workinghours = number_format($row2['workinghours'] / 8, 4);

        // 初始化該日期的報表結構
        if (!isset($construction_report_grouped[$dispatch_date])) {
            $construction_report_grouped[$dispatch_date] = [
                'dispatch_day' => $dispatch_date,
                'support' => ['team_names' => [], 'manpower' => 0, 'workinghours' => 0],
                'supported' => ['team_names' => [], 'manpower' => 0, 'workinghours' => 0],
                'daily_total' => ['manpower' => 0, 'workinghours' => 0],
            ];
        }
        // 根據出勤狀態（派工或被支援）分類資料
        if ($attendance_status == '派工') {
            $construction_report_grouped[$dispatch_date]['support']['team_names'][] = $team_name;
            $construction_report_grouped[$dispatch_date]['support']['manpower'] += $manpower;
            $construction_report_grouped[$dispatch_date]['support']['workinghours'] += $workinghours;
        } else {
            $construction_report_grouped[$dispatch_date]['supported']['team_names'][] = $team_name;
            $construction_report_grouped[$dispatch_date]['supported']['manpower'] += $manpower;
            $construction_report_grouped[$dispatch_date]['supported']['workinghours'] += $workinghours;
        }
        // 累加該日期的總人力與工時
        $construction_report_grouped[$dispatch_date]['daily_total']['manpower'] += $manpower;
        $construction_report_grouped[$dispatch_date]['daily_total']['workinghours'] += $workinghours;

        // 累加進每日總統計（跨工地）
        if (!isset($daily_totals_summary[$dispatch_date])) {
            $daily_totals_summary[$dispatch_date] = ['manpower' => 0, 'workinghours' => 0];
        }
        $daily_totals_summary[$dispatch_date]['manpower'] += $manpower;
        $daily_totals_summary[$dispatch_date]['workinghours'] += $workinghours;
    }

    // 輸出單一工地的統計報表至試算表
    construction_report($start_date, $end_date, $construction_report_grouped, $sheet, $row_offset);

    // 更新行偏移量，為下一個工地報表預留空間
    $row_offset += 6;
}

// 計算最終報表起始行
$final_row = $row_offset + 4;

// 補齊日期區間內無資料的日期，填入 0
$period = new DatePeriod(
    new DateTime($start_date),
    new DateInterval('P1D'),
    (new DateTime($end_date))->modify('+1 day')
);
foreach ($period as $date) {
    $date_str = $date->format('Y-m-d');
    if (!isset($daily_totals_summary[$date_str])) {
        $daily_totals_summary[$date_str] = ['manpower' => 0, 'workinghours' => 0];
    }
}

// 輸出每日總人力與工時統計至試算表
outputDailyTotalsSummary($daily_totals_summary, $sheet, $final_row);
// 輸出公司與工地資訊至試算表
construction_info($construction_sites, $sheet);

// construction_report($start_date, $end_date, $construction_report_grouped, $sheet);

$mDB2->remove();
$mDB->remove();



// 全部天數
function Total_days($start_date, $end_date)
{
	$start_timestamp = strtotime($start_date);
	$end_timestamp = strtotime($end_date);
	$days_difference = ceil(($end_timestamp - $start_timestamp) / (60 * 60 * 24)) + 1; // 加 1 包含開始日
	return $days_difference;
}

// 標題
function table_title($name, $start_date, $end_date, $days, $sheet)
{
	$columnLetter = \PHPExcel_Cell::stringFromColumnIndex(10 + $days);
	$sheet->mergeCells('A1:' . $columnLetter . '2');
	$sheet->setCellValue('A1', $name . $start_date . "~" . $end_date . '工地狀況查詢');

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
	$sheet->getStyle('A1')->getFont()->setSize(20)->setBold(true);
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
	// 因為每個日期佔 4 欄，從第 2 欄開始，所以合計欄是 2 + ($days * 4)
	$columnIndex = 2 + ($days * 4);
	$columnLetter = \PHPExcel_Cell::stringFromColumnIndex($columnIndex);

	// 設定文字置中
	$sheet->getStyle($columnLetter . '3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$sheet->getStyle($columnLetter . '3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

	// 設定底色
	$sheet->getStyle($columnLetter . '3')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
	$sheet->getStyle($columnLetter . '3')->getFill()->getStartColor()->setRGB('7FDBFF');

	// 設定粗體字
	$sheet->getStyle($columnLetter . '3')->getFont()->setBold(true);

	// 設定邊框
	$styleBorders = [
		'borders' => [
			'allborders' => [
				'style' => PHPExcel_Style_Border::BORDER_THIN,
				'color' => ['argb' => 'FF000000'],
			],
		],
	];
	$sheet->getStyle($columnLetter . '3')->applyFromArray($styleBorders);

	// 設定儲存格內容為「合計」
	$sheet->setCellValue($columnLetter . '3', '合計');

	// 設定列寬
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
		// 計算起始欄位索引
		$startColIndex = 2 + ($i * 4);
		$startColLetter = PHPExcel_Cell::stringFromColumnIndex($startColIndex);
		$endColLetter = PHPExcel_Cell::stringFromColumnIndex($startColIndex + 3); // 橫向合併4欄

		// 計算日期
		$current_date = (clone $date)->modify("+$i days")->format('d');

		// 合併儲存格（橫向4格）
		$mergeRange = $startColLetter . '3:' . $endColLetter . '3';
		$sheet->mergeCells($mergeRange);

		// 設定日期到合併儲存格的左上角
		$sheet->setCellValue($startColLetter . '3', $current_date);

		// 設定文字置中
		$sheet->getStyle($mergeRange)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$sheet->getStyle($mergeRange)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

		// 設定底色
		$sheet->getStyle($mergeRange)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
		$sheet->getStyle($mergeRange)->getFill()->getStartColor()->setRGB('7FDBFF');

		// 設定邊框（細線）
		$styleBorders = [
			'borders' => [
				'allborders' => [
					'style' => PHPExcel_Style_Border::BORDER_THIN,
					'color' => ['argb' => 'FF000000'],
				],
			],
		];
		$sheet->getStyle($mergeRange)->applyFromArray($styleBorders);
	}
}

//公司資訊 
function construction_info($members, $sheet, $startRow = 4)
{
	$rowsPerMember = 6;
	$memberCount = count($members);

	for ($i = 0; $i < $memberCount; $i++) {
		$currentRow = $startRow + ($i * $rowsPerMember);

		$sheet->mergeCells("A{$currentRow}:A" . ($currentRow + $rowsPerMember - 1));
		$sheet->setCellValue("A{$currentRow}", $i + 1);

		$sheet->mergeCells("B{$currentRow}:B" . ($currentRow + $rowsPerMember - 1));
		$sheet->setCellValue("B{$currentRow}", $members[$i]['construction_site'] . "\n" . $members[$i]['construction_id']);
		$sheet->getStyle("B{$currentRow}")
			->getAlignment()->setWrapText(true);

		$cellRange = "A{$currentRow}:B" . ($currentRow + $rowsPerMember - 1);
		$sheet->getStyle($cellRange)
			->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		$sheet->getStyle($cellRange)
			->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
		$sheet->getStyle($cellRange)
			->getFont()->setSize(12);
		$sheet->getStyle($cellRange)->applyFromArray([
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

	// 合計列（A欄垂直合併兩列，B欄第一列為總人數，第二列為總工數）
	$totalRow = $startRow + ($memberCount * $rowsPerMember);
	$nextRow = $totalRow + 1;

	// A欄合併兩列，顯示「合計」
	$sheet->mergeCells("A{$totalRow}:A{$nextRow}");
	$sheet->setCellValue("A{$totalRow}", "合計");

	// B欄第一列為「總人數：X」
	$sheet->setCellValue("B{$totalRow}", "總人數");

	// 計算總工數，假設這是根據每位成員的工數來計算的
	$totalWorkCount = $memberCount * 5;  // 假設每位成員有5工數，根據實際需求調整

	// B欄第二列為「總工數：X」
	$sheet->setCellValue("B{$nextRow}", "總工數");

	// 套用樣式
	$styleRange = "A{$totalRow}:B{$nextRow}";
	$sheet->getStyle($styleRange)
		->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$sheet->getStyle($styleRange)
		->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$sheet->getStyle($styleRange)
		->getFont()->setBold(true)
		->setSize(12);
	$sheet->getStyle($styleRange)->applyFromArray([
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

	// 設定底色
	$sheet->getStyle($styleRange)
		->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
	$sheet->getStyle($styleRange)
		->getFill()->getStartColor()->setRGB('FFDC00');

}

// 報表輸出
function construction_report($start_date, $end_date, $construction_report_grouped, $sheet, $row_offset = 0)
{
    // 檢查工作表物件是否有效
    if ($sheet === null) {
        throw new Exception("無效的工作表物件。");
    }

    // 計算總天數
    $total_days = Total_days($start_date, $end_date);
    $column_offset = 2; // 從 C 欄開始寫入資料
    $block_height = 6;  // 每筆工地資料佔用的行數
    $initial_row = 3 + $row_offset; // 起始行數

    $current_column_index = $column_offset; // 當前欄位索引
    $current_date = $start_date; // 當前日期

    // 初始化橫向總計變數
    $grand_total_manpower = 0; // 總人力
    $grand_total_workinghours = 0; // 總工時

    // 遍歷每個日期
    for ($day = 0; $day <= $total_days; $day++) {

        $total_manpower = 0; // 當日總人力
        $total_workinghours = 0; // 當日總工時
        $column_letter = \PHPExcel_Cell::stringFromColumnIndex($current_column_index); // 當前欄位字母

        // 檢查當日是否有報表資料
        if (isset($construction_report_grouped[$current_date])) {
            $report = $construction_report_grouped[$current_date];

            // 計算合併欄位的起始和結束索引
            $startColumnIndex = \PHPExcel_Cell::columnIndexFromString($column_letter) - 1;
            $endColumnIndex = $startColumnIndex + 3;
            $endColumnLetter = \PHPExcel_Cell::stringFromColumnIndex($endColumnIndex);

            $prefix_symbol = '➤'; // 用於格式化小隊名稱的前綴符號

            // 合併第二行並寫入派工小隊資料
            $mergeRow2 = $initial_row + 1;
            $sheet->mergeCells($column_letter . $mergeRow2 . ':' . $endColumnLetter . $mergeRow2);

            // 格式化派工小隊名稱
            $support_team_lines = array_map(function ($team) use ($prefix_symbol) {
                return $prefix_symbol . ' ' . $team;
            }, $report['support']['team_names']);

            // 設置派工小隊文字
            $support_text = "派工小隊\n" . implode("\n", $support_team_lines);
            $sheet->setCellValue($column_letter . $mergeRow2, $support_text);

            // 設置派工小隊單元格樣式
            $style = $sheet->getStyle($column_letter . $mergeRow2);
            $style->getAlignment()->setWrapText(true); // 自動換行
            $style->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_TOP); // 垂直置頂
            $style->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER); // 水平置中

            // 根據小隊數量動態調整行高
            $support_team_count = count($report['support']['team_names']);
            $sheet->getRowDimension($mergeRow2)->setRowHeight(50 * ($support_team_count + 1));

            // 寫入派工小隊的人力和工時資料
            $sheet->setCellValue(\PHPExcel_Cell::stringFromColumnIndex($current_column_index) . ($initial_row + 2), '人');
            $sheet->setCellValue(\PHPExcel_Cell::stringFromColumnIndex($current_column_index + 1) . ($initial_row + 2), $report['support']['manpower']);
            $sheet->setCellValue(\PHPExcel_Cell::stringFromColumnIndex($current_column_index + 2) . ($initial_row + 2), '工時');
            $sheet->setCellValue(\PHPExcel_Cell::stringFromColumnIndex($current_column_index + 3) . ($initial_row + 2), $report['support']['workinghours']);

            // 合併第三行並寫入被支援小隊資料
            $mergeRow3 = $initial_row + 3;
            $sheet->mergeCells($column_letter . $mergeRow3 . ':' . $endColumnLetter . $mergeRow3);

            // 格式化被支援小隊名稱
            $supported_team_lines = array_map(function ($team) use ($prefix_symbol) {
                return $prefix_symbol . ' ' . $team;
            }, $report['supported']['team_names']);

            // 設置被支援小隊文字
            $supported_text = "被支援小隊\n" . implode("\n", $supported_team_lines);
            $sheet->setCellValue($column_letter . $mergeRow3, $supported_text);

            // 設置被支援小隊單元格樣式
            $style = $sheet->getStyle($column_letter . $mergeRow3);
            $style->getAlignment()->setWrapText(true); // 自動換行
            $style->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_TOP); // 垂直置頂
            $style->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER); // 水平置中

            // 根據小隊數量動態調整行高
            $supported_team_count = count($report['supported']['team_names']);
            $sheet->getRowDimension($mergeRow3)->setRowHeight(50 * ($supported_team_count + 1));

            // 寫入被支援小隊的人力和工時資料
            $sheet->setCellValue(\PHPExcel_Cell::stringFromColumnIndex($current_column_index) . ($initial_row + 4), '人');
            $sheet->setCellValue(\PHPExcel_Cell::stringFromColumnIndex($current_column_index + 1) . ($initial_row + 4), $report['supported']['manpower']);
            $sheet->setCellValue(\PHPExcel_Cell::stringFromColumnIndex($current_column_index + 2) . ($initial_row + 4), '工時');
            $sheet->setCellValue(\PHPExcel_Cell::stringFromColumnIndex($current_column_index + 3) . ($initial_row + 4), $report['supported']['workinghours']);

            // 計算當日總人力和總工時
            $total_manpower = $report['support']['manpower'] + $report['supported']['manpower'];
            $total_workinghours = $report['support']['workinghours'] + $report['supported']['workinghours'];

            // 合併第五行並寫入派工合計標題
            $mergeRow1 = $initial_row + 5;
            $mergeRange = $column_letter . $mergeRow1 . ':' . $endColumnLetter . $mergeRow1;
            $sheet->mergeCells($mergeRange);
            $sheet->setCellValue($column_letter . $mergeRow1, '派工合計');
            $sheet->getStyle($mergeRange)->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER); // 水平置中
            $sheet->getStyle($mergeRange)->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER); // 垂直置中
            $sheet->getRowDimension($mergeRow1)->setRowHeight(30); // 設置行高

            // 寫入合計的人力和工時資料
            $sheet->setCellValue(\PHPExcel_Cell::stringFromColumnIndex($current_column_index) . ($initial_row + 6), '人');
            $sheet->setCellValue(\PHPExcel_Cell::stringFromColumnIndex($current_column_index + 1) . ($initial_row + 6), $total_manpower);
            $sheet->setCellValue(\PHPExcel_Cell::stringFromColumnIndex($current_column_index + 2) . ($initial_row + 6), '工時');
            $sheet->setCellValue(\PHPExcel_Cell::stringFromColumnIndex($current_column_index + 3) . ($initial_row + 6), $total_workinghours);

            // 累加橫向總計
            $grand_total_manpower += $total_manpower;
            $grand_total_workinghours += $total_workinghours;
        }

        // 日期遞增並更新欄位索引
        $current_date = date('Y-m-d', strtotime($current_date . ' +1 day'));
        $current_column_index += 4;
    }

    // 設定所有欄位的寬度
    for ($i = $column_offset; $i < $current_column_index; $i++) {
        $colLetter = \PHPExcel_Cell::stringFromColumnIndex($i);
        $sheet->getColumnDimension($colLetter)->setWidth(5);
    }

    // 設置整個資料範圍的邊框和對齊樣式
    $end_column_letter = \PHPExcel_Cell::stringFromColumnIndex($current_column_index - 5);
    $end_row = $initial_row + $block_height;
    $full_range = "C" . ($initial_row + 1) . ":" . $end_column_letter . $end_row;

    $sheet->getStyle($full_range)->applyFromArray([
        'borders' => [
            'allborders' => [
                'style' => \PHPExcel_Style_Border::BORDER_THIN, // 細邊框
                'color' => ['rgb' => '000000'],
            ],
        ],
        'alignment' => [
            'horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER, // 水平置中
            'wrap' => true, // 自動換行
        ],
    ]);

    // 為第五、第六行設置粗邊框
    $bold_start_row = $initial_row + 5;
    $bold_end_row = $bold_start_row + 1;

    for ($i = $column_offset; $i < $current_column_index - 4; $i += 4) {
        $startCol = \PHPExcel_Cell::stringFromColumnIndex($i);
        $endCol = \PHPExcel_Cell::stringFromColumnIndex($i + 3);
        $bold_range = $startCol . $bold_start_row . ':' . $endCol . $bold_end_row;

        $sheet->getStyle($bold_range)->applyFromArray([
            'borders' => [
                'allborders' => [
                    'style' => \PHPExcel_Style_Border::BORDER_MEDIUM, // 粗邊框
                    'color' => ['rgb' => '000000'],
                ],
            ],
        ]);
    }

    // 在最後一欄顯示橫向總計
    $column_letter = \PHPExcel_Cell::stringFromColumnIndex($current_column_index - 4); // 使用最後一組資料的起始欄位
    $row_start = $initial_row + 1;
    $row_end = $row_start + 5; // 合併 6 列
    $merge_range = $column_letter . $row_start . ':' . $column_letter . $row_end;

    $sheet->mergeCells($merge_range);

    // 設置總計文字（總人數和總工時）
    $summary_text = '總人數: ' . $grand_total_manpower . "\n" . '總工時: ' . $grand_total_workinghours;
    $sheet->setCellValue($column_letter . $row_start, $summary_text);

    // 設置總計單元格樣式
    $style = $sheet->getStyle($merge_range);
    $style->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER); // 水平置中
    $style->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER); // 垂直置中
    $style->getAlignment()->setWrapText(true); // 自動換行
    $style->getFont()->setBold(true); // 粗體

    // 設置總計區域的邊框
    $style->getBorders()->getAllBorders()->setBorderStyle(\PHPExcel_Style_Border::BORDER_THIN);
    $style->getBorders()->getAllBorders()->getColor()->setRGB('000000');

    // 設置總計欄寬
    $sheet->getColumnDimension($column_letter)->setWidth(10);
}

// 建立每日總計
function outputDailyTotalsSummary($daily_totals_summary, $sheet, &$start_row) {
    ksort($daily_totals_summary);
    
    // 設定欄位起始位置
    $start_col = 'C';
    $col_index = 2; // C 欄對應索引 2 (A=0, B=1, C=2)
    
    // 初始化總計變數
    $total_manpower = 0;
    $total_workinghours = 0;
    
    // 定義樣式
    $styleArray = [
        'alignment' => [
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'allborders' => [
                'style' => PHPExcel_Style_Border::BORDER_THICK,
                'color' => ['rgb' => '000000'],
            ],
        ],
        'font' => [
            'bold' => true,
        ],
        'fill' => [
            'type' => PHPExcel_Style_Fill::FILL_SOLID,
            'color' => ['rgb' => 'FFDC00'],
        ],
    ];
    
    // 遍歷資料
    foreach ($daily_totals_summary as $totals) {
        // 計算合併的結束欄位（當前欄位 + 3）
        $end_col = PHPExcel_Cell::stringFromColumnIndex($col_index + 3);
        
        // 寫入 manpower（第一列）並合併 4 格
        $sheet->setCellValue($start_col . $start_row, $totals['manpower']);
        $sheet->mergeCells($start_col . $start_row . ':' . $end_col . $start_row);
        $sheet->getStyle($start_col . $start_row . ':' . $end_col . $start_row)->applyFromArray($styleArray);
        
        // 寫入 workinghours（第二列，格式化為三位小數）並合併 4 格
        $sheet->setCellValue($start_col . ($start_row + 1), number_format($totals['workinghours'], 3));
        $sheet->mergeCells($start_col . ($start_row + 1) . ':' . $end_col . ($start_row + 1));
        $sheet->getStyle($start_col . ($start_row + 1) . ':' . $end_col . ($start_row + 1))->applyFromArray($styleArray);
        
        // 累加總計
        $total_manpower += $totals['manpower'];
        $total_workinghours += $totals['workinghours'];
        
        // 更新欄位索引（每次跳 4 格）
        $col_index += 4;
        $start_col = PHPExcel_Cell::stringFromColumnIndex($col_index);
    }
    
    // 寫入總計到下一欄（不合併）
    // 寫入 total_manpower（第一列）
    $sheet->setCellValue($start_col . $start_row, $total_manpower);
    $sheet->getStyle($start_col . $start_row)->applyFromArray($styleArray);
    
    // 寫入 total_workinghours（第二列，格式化為三位小數）
    $sheet->setCellValue($start_col . ($start_row + 1), number_format($total_workinghours, 3));
    $sheet->getStyle($start_col . ($start_row + 1))->applyFromArray($styleArray);
}


// Rename worksheet
$objPHPExcel->getActiveSheet()->setTitle("工地狀況查詢");


$xlsx_filename = $company_name . "_" . $start_date . "~" . $end_date . "_工地狀況查詢.xls";


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

?>