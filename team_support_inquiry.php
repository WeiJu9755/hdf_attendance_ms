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


$m_location		= "/website/smarty/templates/".$site_db."/".$templates;
$m_pub_modal	= "/website/smarty/templates/".$site_db."/pub_modal";


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

@include_once("/website/class/".$site_db."_info_class.php");



$fm = $_GET['fm'];
$ch = $_GET['ch'];

$team_id = $_GET['team_id'];
$team_id2 = $_GET['team_id2'];
$team_construction_id = $_GET['team_construction_id'];


$start_date = $_GET['start_date'];
$end_date = $_GET['end_date'];

if (!isset($_GET['start_date']))
	$start_date = date('Y-m-d');

if (!isset($_GET['end_date']))
	$end_date = date('Y-m-d');


//檢查是否為管理員及進階會員
$super_admin = "N";
$super_advanced = "N";
$mem_row = getkeyvalue2('memberinfo','member',"member_no = '$memberID'",'admin,advanced');
$super_admin = $mem_row['admin'];
$super_advanced = $mem_row['advanced'];



$mDB = "";
$mDB = new MywebDB();

$mDB2 = "";
$mDB2 = new MywebDB();


//載入團隊
if  ((($super_advanced=="Y") && ($advanced_readonly <> "Y")) && ($super_admin <> "Y")) {
	$Qry="SELECT a.team_id,a.team_name FROM team a
	LEFT JOIN company b ON b.company_id = a.company_id
	RIGHT JOIN group_company c ON c.company_id = a.company_id and c.member_no = '$memberID'
	WHERE a.team_id <> ''
	ORDER BY a.team_id";
} else {
	$Qry="SELECT team_id,team_name FROM team ORDER BY team_id";
}
$mDB->query($Qry);

$select_team = "";
$select_team  = "<select class=\"inline form-select\" name=\"team_list\" id=\"team_list\" style=\"width:auto;\">";
$select_team .= "<option></option>";
if ($mDB->rowCount() > 0) {
	while ($row=$mDB->fetchRow(2)) {
		$ch_team_id = $row['team_id'];
		$ch_team_name = $row['team_name'];
		$select_team .= "<option value='$ch_team_id' ".mySelect($ch_team_id,$team_id).">$ch_team_name $ch_team_id</option>";
	}
}
$select_team .= "</select>";

//載入支援團隊
/*
if  ((($super_advanced=="Y") && ($advanced_readonly <> "Y")) && ($super_admin <> "Y")) {
	$Qry="SELECT a.team_id,a.team_name FROM team a
	LEFT JOIN company b ON b.company_id = a.company_id
	RIGHT JOIN group_company c ON c.company_id = a.company_id and c.member_no = '$memberID'
	WHERE a.team_id <> ''
	ORDER BY a.team_id";
} else {
	$Qry="SELECT team_id,team_name FROM team ORDER BY team_id";
}
*/
$Qry="SELECT team_id,team_name FROM team ORDER BY team_id";
$mDB->query($Qry);

$select_team2 = "";
$select_team2  = "<select class=\"inline form-select\" name=\"team2_list\" id=\"team2_list\" style=\"width:auto;\">";
$select_team2 .= "<option></option>";
if ($mDB->rowCount() > 0) {
	while ($row=$mDB->fetchRow(2)) {
		$ch_support_team_id = $row['team_id'];
		$ch_support_team_name = $row['team_name'];
		$select_team2 .= "<option value='$ch_support_team_id' ".mySelect($ch_support_team_id,$team_id2).">$ch_support_team_name $ch_support_team_id</option>";
	}
}
$select_team2 .= "</select>";

//載入團隊工地
$Qry="select construction_id,construction_site from construction order by auto_seq";
$mDB->query($Qry);

$select_team_construction = "";
$select_team_construction  = "<select class=\"inline form-select\" name=\"team_construction_list\" id=\"team_construction_list\" style=\"width:auto;\">";
$select_team_construction .= "<option></option>";

if ($mDB->rowCount() > 0) {
	while ($row=$mDB->fetchRow(2)) {
		$ch_construction_id = $row['construction_id'];
		$ch_construction_site = $row['construction_site'];
		$select_team_construction .= "<option value=\"$ch_construction_id\" ".mySelect($ch_construction_id,$team_construction_id).">$ch_construction_id $ch_construction_site</option>";
	}
}
$select_team_construction .= "</select>";




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
$m_attendance_day = [];
$m_attendance_str = [];

// 將日期加入陣列
foreach ($period as $date) {
    $m_header[] = [$date->format('Y-m-d'),$date->format('d')];
    $m_attendance_day_TOT[] = 0;
	$m_attendance_day[] = 0;
	$m_attendance_str[] = "";
}

// 輸出日期陣列
//print_r($dateArray);
//exit;


$show_disabled = "style=\"pointer-events: none;\"";

$show_inquiry = "";


//取得工單資料
/*
if (!empty($team_construction_id)) {
	if (!empty($team_id2)) {
		$Qry="SELECT a.dispatch_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day,b.employee_id,c.employee_name,a.team_id,d.team_name 
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
		$Qry="SELECT a.dispatch_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day,b.employee_id,c.employee_name,a.team_id,d.team_name 
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
		$Qry="SELECT a.dispatch_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day,b.employee_id,c.employee_name,a.team_id,d.team_name 
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
		$Qry="SELECT a.dispatch_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day,b.employee_id,c.employee_name,a.team_id,d.team_name 
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
*/

if (!empty($team_id)) {

	if (!empty($team_construction_id)) {
		if (!empty($team_id2)) {
			$Qry="SELECT b.employee_id,c.employee_name,a.team_id,d.team_name 
			FROM dispatch a
			LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
			LEFT JOIN employee c ON c.employee_id = b.employee_id
			LEFT JOIN team d ON d.team_id = a.team_id
			LEFT JOIN team e ON e.team_id = b.team_id
			LEFT JOIN construction f ON f.construction_id = b.team_construction_id
			WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' 
			AND a.team_id = '$team_id' 
			AND (b.team_id = '$team_id2' OR (b.transition = 'Y' AND b.transition_team_id = '$team_id2'))
			GROUP BY b.employee_id
			ORDER BY c.team_id,b.employee_id";
		} else {
			$Qry="SELECT b.employee_id,c.employee_name,a.team_id,d.team_name 
		
			FROM dispatch a
			LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
			LEFT JOIN employee c ON c.employee_id = b.employee_id
			LEFT JOIN team d ON d.team_id = a.team_id
			LEFT JOIN team e ON e.team_id = b.team_id
			LEFT JOIN construction f ON f.construction_id = b.team_construction_id
			WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' 
			AND a.team_id = '$team_id' 
			AND b.team_construction_id = '$team_construction_id'
			GROUP BY b.employee_id
			ORDER BY c.team_id,b.employee_id";
		}
	} else {
		if (!empty($team_id2)) {
			$Qry="SELECT b.employee_id,c.employee_name,a.team_id,d.team_name 
			FROM dispatch a
			LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
			LEFT JOIN employee c ON c.employee_id = b.employee_id
			LEFT JOIN team d ON d.team_id = a.team_id
			LEFT JOIN team e ON e.team_id = b.team_id
			LEFT JOIN construction f ON f.construction_id = b.team_construction_id
			WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' 
			AND a.team_id = '$team_id' 
			AND (b.team_id = '$team_id2' OR (b.transition = 'Y' AND b.transition_team_id = '$team_id2'))
			GROUP BY b.employee_id
			ORDER BY c.team_id,b.employee_id";
		} else {
			$Qry="SELECT b.employee_id,c.employee_name,a.team_id,d.team_name 
			FROM dispatch a
			LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
			LEFT JOIN employee c ON c.employee_id = b.employee_id
			LEFT JOIN team d ON d.team_id = a.team_id
			LEFT JOIN team e ON e.team_id = b.team_id
			LEFT JOIN construction f ON f.construction_id = b.team_construction_id
			WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' 
			AND a.team_id = '$team_id' 
			GROUP BY b.employee_id
			ORDER BY c.team_id,b.employee_id";
		}
	}

} else {

	if (!empty($team_construction_id)) {
		if (!empty($team_id2)) {
			$Qry="SELECT b.employee_id,c.employee_name,a.team_id,d.team_name 
			FROM dispatch a
			LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
			LEFT JOIN employee c ON c.employee_id = b.employee_id
			LEFT JOIN team d ON d.team_id = a.team_id
			LEFT JOIN team e ON e.team_id = b.team_id
			LEFT JOIN construction f ON f.construction_id = b.team_construction_id
			WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' 
			AND (b.team_id = '$team_id2' OR (b.transition = 'Y' AND b.transition_team_id = '$team_id2'))
			GROUP BY b.employee_id
			ORDER BY c.team_id,b.employee_id";
		} else {
			$Qry="SELECT b.employee_id,c.employee_name,a.team_id,d.team_name 
		
			FROM dispatch a
			LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
			LEFT JOIN employee c ON c.employee_id = b.employee_id
			LEFT JOIN team d ON d.team_id = a.team_id
			LEFT JOIN team e ON e.team_id = b.team_id
			LEFT JOIN construction f ON f.construction_id = b.team_construction_id
			WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' 
			AND b.team_construction_id = '$team_construction_id'
			GROUP BY b.employee_id
			ORDER BY c.team_id,b.employee_id";
		}
	} else {
		if (!empty($team_id2)) {
			$Qry="SELECT b.employee_id,c.employee_name,a.team_id,d.team_name 
			FROM dispatch a
			LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
			LEFT JOIN employee c ON c.employee_id = b.employee_id
			LEFT JOIN team d ON d.team_id = a.team_id
			LEFT JOIN team e ON e.team_id = b.team_id
			LEFT JOIN construction f ON f.construction_id = b.team_construction_id
			WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' 
			AND (b.team_id = '$team_id2' OR (b.transition = 'Y' AND b.transition_team_id = '$team_id2'))
			GROUP BY b.employee_id
			ORDER BY c.team_id,b.employee_id";
		} else {
			$Qry="SELECT b.employee_id,c.employee_name,a.team_id,d.team_name 
			FROM dispatch a
			LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
			LEFT JOIN employee c ON c.employee_id = b.employee_id
			LEFT JOIN team d ON d.team_id = a.team_id
			LEFT JOIN team e ON e.team_id = b.team_id
			LEFT JOIN construction f ON f.construction_id = b.team_construction_id
			WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' 
			GROUP BY b.employee_id
			ORDER BY c.team_id,b.employee_id";
		}
	}

}

//echo $Qry;
//exit; 

$mDB->query($Qry);


$attendance_day_TOTAL = 0;

$alist_kwh = array();


if ($mDB->rowCount() > 0) {

	if ($powerkey=="A") {
		$show_disabled = "";
	}

	//顯示抬頭標題列
$show_inquiry.=<<<EOT
	<table class="table table-bordered" style="border: 2px solid #000;background-color: #FFFFFF;">
		<thead>
			<tr class="text-center" style="border-bottom: 2px solid #000;">
				<th scope="col" class="size12 bg-aqua text-nowrap vmiddle" style="padding: 10px;width:2%;background-color: #D2F2FF;"><b>序號</b></th>
				<th scope="col" class="size12 bg-aqua text-nowrap vmiddle" style="padding: 10px;width:5%;background-color: #D2F2FF;"><b>團隊</b></th>
				<th scope="col" class="size12 bg-aqua text-nowrap vmiddle" style="padding: 10px;width:5%;background-color: #D2F2FF;"><b>員工</b></th>
EOT;

$i_count = count($m_header);
for ($i = 0; $i < $i_count; $i++) {
    $show_inquiry.="<th scope=\"col\" class=\"size14 bg-aqua vmiddle\" style=\"width:3%;padding: 10px;background-color: #D2F2FF;\"><b>".$m_header[$i][1]."</b></th>";
}

//<th scope="col" class="size14 bg-aqua" style="padding: 10px 0;width:150px;background-color: #D2F2FF;"><b>合計</b></th>

$show_inquiry.=<<<EOT
				<th scope="col" class="size14 bg-aqua" style="padding: 10px;width:5%;background-color: #D2F2FF;"><b>合計</b></th>
			</tr>
		</thead>
		<tbody>
EOT;

	$seq = 0;
	while ($row=$mDB->fetchRow(2)) {
		//$dispatch_id = $row['dispatch_id'];
		//$dispatch_year = $row['dispatch_year'];
		//$dispatch_month = $row['dispatch_month'];
		//$dispatch_day = $row['dispatch_day'];
		$employee_id = $row['employee_id'];
		$employee_name = $row['employee_name'];
		$team_id = $row['team_id'];
		$team_name = $row['team_name'];

		//$team_id2 = $row['team_id2'];
		//$team_name2 = $row['team_name2'];

		//$team_construction_id = $row['team_construction_id'];
		//$construction_site = $row['construction_site'];

		if (!empty($employee_id)) {
			$seq++;

			//再取得各員工的資料
			if (!empty($team_construction_id)) {
				if (!empty($team_id2)) {
					$Qry2="SELECT a.dispatch_id,a.dispatch_date,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day
						,b.employee_id,b.manpower,b.attendance_day,b.attendance_status,b.transition,b.transition_start,b.transition_end
						,b.team_id as support_team_id,c.team_name as team_name2,b.transition_team_id,d.team_name as transition_team_name,b.team_construction_id
						,f.construction_site,g.construction_site  as transition_construction_site
						FROM dispatch a
						LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
						LEFT JOIN team c ON c.team_id = b.team_id
						LEFT JOIN team d ON d.team_id = b.transition_team_id
						LEFT JOIN construction f ON f.construction_id = b.team_construction_id
						LEFT JOIN construction g ON g.construction_id = b.transition_construction_id
						WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' AND b.employee_id = '$employee_id'
						AND a.team_id = '$team_id' 
						AND (b.team_id = '$team_id2' OR (b.transition = 'Y' AND b.transition_team_id = '$team_id2'))
						AND b.team_construction_id = '$team_construction_id'
						ORDER BY b.employee_id,a.dispatch_date";
				} else {
					$Qry2="SELECT a.dispatch_id,a.dispatch_date,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day
						,b.employee_id,b.manpower,b.attendance_day,b.attendance_status,b.transition,b.transition_start,b.transition_end
						,b.team_id as support_team_id,c.team_name as team_name2,b.transition_team_id,d.team_name as transition_team_name,b.team_construction_id
						,f.construction_site,g.construction_site  as transition_construction_site
						FROM dispatch a
						LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
						LEFT JOIN team c ON c.team_id = b.team_id
						LEFT JOIN team d ON d.team_id = b.transition_team_id
						LEFT JOIN construction f ON f.construction_id = b.team_construction_id
						LEFT JOIN construction g ON g.construction_id = b.transition_construction_id
						WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' AND b.employee_id = '$employee_id'
						AND a.team_id = '$team_id'
						AND b.team_construction_id = '$team_construction_id'
						ORDER BY b.employee_id,a.dispatch_date";
				}
			} else {
				if (!empty($team_id2)) {
					$Qry2="SELECT a.dispatch_id,a.dispatch_date,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day
						,b.employee_id,b.manpower,b.attendance_day,b.attendance_status,b.transition,b.transition_start,b.transition_end
						,b.team_id as support_team_id,c.team_name as team_name2,b.transition_team_id,d.team_name as transition_team_name,b.team_construction_id
						,f.construction_site,g.construction_site  as transition_construction_site
						FROM dispatch a
						LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
						LEFT JOIN team c ON c.team_id = b.team_id
						LEFT JOIN team d ON d.team_id = b.transition_team_id
						LEFT JOIN construction f ON f.construction_id = b.team_construction_id
						LEFT JOIN construction g ON g.construction_id = b.transition_construction_id
						WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' AND b.employee_id = '$employee_id'
						AND a.team_id = '$team_id' 
						AND (b.team_id = '$team_id2' OR (b.transition = 'Y' AND b.transition_team_id = '$team_id2'))
						ORDER BY b.employee_id,a.dispatch_date";
				} else {
					$Qry2="SELECT a.dispatch_id,a.dispatch_date,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day
						,b.employee_id,b.manpower,b.attendance_day,b.attendance_status,b.transition,b.transition_start,b.transition_end
						,b.team_id as support_team_id,c.team_name as team_name2,b.transition_team_id,d.team_name as transition_team_name,b.team_construction_id
						,f.construction_site,g.construction_site  as transition_construction_site
						FROM dispatch a
						LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
						LEFT JOIN team c ON c.team_id = b.team_id
						LEFT JOIN team d ON d.team_id = b.transition_team_id
						LEFT JOIN construction f ON f.construction_id = b.team_construction_id
						LEFT JOIN construction g ON g.construction_id = b.transition_construction_id
						WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' AND b.employee_id = '$employee_id'
						AND a.team_id = '$team_id'
						ORDER BY b.employee_id,a.dispatch_date";
				}
			}

			//echo $Qry2;
			//echo "<br>";
			//exit;

			$mDB2->query($Qry2);
			$SUMMARY = 0;
			if ($mDB2->rowCount() > 0) {

				$m_attendance_day = [];
				foreach ($period as $date) {
					$m_attendance_day[] = "&nbsp;";
				}
				$m_attendance_str = [];
				foreach ($period as $date) {
					$m_attendance_str[] = "&nbsp;";
				}
				$m_support = [];
				foreach ($period as $date) {
					$m_support[] = "&nbsp;";
				}

				while ($row2=$mDB2->fetchRow(2)) {

					$support_team_id = $row2['support_team_id'];
					$team_name2 = $row2['team_name2'];

					//$team_construction_id = $row2['team_construction_id'];

					//支援工地
					$construction_site = $row2['construction_site'];

					//轉場工地
					$transition_construction_site = $row2['transition_construction_site'];
			
					$attendance_day = $row2['attendance_day'];
					
					$DAY = $row2['dispatch_date'];


					//echo  $employee_id." ".$team_id2." ".$DAY."<br>";

					$i_count = count($m_attendance_day);
					for ($i = 0; $i < $i_count; $i++) {

						if ($DAY == $m_header[$i][0]) {

							
							$temp_str = "";
							$status_str = "";
							$transition_hours = 0;

							//轉場
							if ($row2['transition'] == "Y") {

								//計算轉場時數/日
								$transition_start = $row2['transition_start'];
								$transition_end = $row2['transition_end'];

								$transition_hours = round(calculateLeaveHours($transition_start, $transition_end)/8,4);

								$status_str .= "<span class=\"text-nowrap\">轉場 ".$row2['transition_team_name'].$transition_hours."</span>";

								if (!empty($transition_construction_site))							
									$status_str .= "<div class=\"text-nowrap\">".$transition_construction_site."</div>";

								$m_support[$i] = $status_str;

								$SUMMARY = $SUMMARY+$transition_hours;	//橫向加總

								//累計加總
								$m_attendance_day_TOT[$i] = $m_attendance_day_TOT[$i]+$transition_hours;


							}


							//支援
							if ($row2['attendance_status'] == "支援") {


								if ($transition_hours <> 0) {
									$support_hours = round($attendance_day/8,4)-$transition_hours;
								} else {
									$support_hours = round($attendance_day/8,4);
								}

								if (!empty($team_id2)) {

									if ($support_team_id == $team_id2) {
										//計算支援時數/日
										//$support_hours = round($attendance_day/8,4);
										$temp_str = "<div class=\"text-nowrap\">".$row2['attendance_status']." ".$team_name2."</div>";
										if (!empty($construction_site))							
											$temp_str .= "<div class=\"text-nowrap\">".$construction_site."</div>";
									}
									$m_attendance_day[$i] = $support_hours;
									$m_attendance_str[$i] = $temp_str;

									$SUMMARY = $SUMMARY+$support_hours;	//橫向加總

									//累計加總
									$m_attendance_day_TOT[$i] = $m_attendance_day_TOT[$i]+$support_hours;

								} else {
								
									//計算支援時數/日
									//$support_hours = round($row2['attendance_day']/8,4);
									$temp_str = "<div class=\"text-nowrap\">".$row2['attendance_status']." ".$team_name2."</div>";
									if (!empty($construction_site))							
										$temp_str .= "<div class=\"text-nowrap\">".$construction_site."</div>";
									$m_attendance_day[$i] = $support_hours;
									$m_attendance_str[$i] = $temp_str;

									$SUMMARY = $SUMMARY+$support_hours;	//橫向加總

									//累計加總
									$m_attendance_day_TOT[$i] = $m_attendance_day_TOT[$i]+$support_hours;

								}

							}


							/*
							$m_attendance_day[$i] = round($row2['attendance_day']/8,4);
							$SUMMARY = $SUMMARY+$m_attendance_day[$i];	//橫向加總
							//累計加總
							$m_attendance_day_TOT[$i] = $m_attendance_day_TOT[$i]+$m_attendance_day[$i];
							*/
							
							
						}
						/*
						if ($m_attendance_day[$i] == 0) {
							$m_attendance_day[$i] = "";
						}
						*/

					}

				}

			}
			$attendance_day_TOTAL = $attendance_day_TOTAL+$SUMMARY;


$show_inquiry.=<<<EOT
			<tr class="text-center">
				<td scope="row" rowspan="3" class="text-nowrap vmiddle">$seq</td>
				<td scope="row" rowspan="3" class="text-nowrap vmiddle"><div class="size14 weight">$team_name</div><div class="size08">$team_id</div></td>
				<td scope="row" rowspan="3" class="text-nowrap vmiddle"><div class="size14 weight">$employee_name</div><div class="size08">$employee_id</div></td>
EOT;

//支援
$i_count = count($m_attendance_day);
for ($i = 0; $i < $i_count; $i++) {
    $show_inquiry.="<td class=\"text-center vmiddle\">".$m_attendance_day[$i]."</td>";
}

$show_inquiry.=<<<EOT
				<td scope="row" rowspan="3" class="text-nowrap vmiddle"><div class="size14 weight">$SUMMARY</div></td>
			</tr>
EOT;		

$show_inquiry.=<<<EOT
			<tr class="text-center">
EOT;

$i_count = count($m_attendance_str);
for ($i = 0; $i < $i_count; $i++) {
    $show_inquiry.="<td class=\"text-center vmiddle\">".$m_attendance_str[$i]."</td>";
}

$show_inquiry.=<<<EOT
			</tr>
			<tr>
EOT;		


//轉場
$i_count = count($m_support);
for ($i = 0; $i < $i_count; $i++) {
    $show_inquiry.="<td class=\"text-center vmiddle\">".$m_support[$i]."</td>";
}

$show_inquiry.=<<<EOT
			</tr>
			<tr>
EOT;		


		}

	}

	//exit;

//顯示全部總和
$show_inquiry.=<<<EOT
   <tr class="text-center bg-yellow size14 weight" style="border-top: 2px solid #000;">
	   <th scope="row" colspan="3" class="text-nowrap text-center" style="background-color: #FFEBAC;">合計</th>
EOT;

$i_count = count($m_attendance_day_TOT);
$attendance_day_all = 0;
for ($i = 0; $i < $i_count; $i++) {
    $show_inquiry.="<td style=\"padding: 10px;background-color: #FFEBAC;\">".$m_attendance_day_TOT[$i]."</td>";
	$attendance_day_all = $attendance_day_all+$m_attendance_day_TOT[$i];
}
$attendance_day_all = round($attendance_day_all,4);

$show_inquiry.=<<<EOT
	   <td class="text-center size12 weight" style="background-color: #FFEBAC;"><i>$attendance_day_all</i></td>
   </tr>
EOT;			


$show_inquiry.=<<<EOT
			</tr>
		</tbody>
	</table>
EOT;


} else {

	$show_inquiry = "<div class=\"size16 weight text-center m-3\">查無任何資料！</div>";

}


$mDB2->remove();
$mDB->remove();




/*
$Close = getlang("關閉");
$Print = getlang("列印");
*/

$show_report=<<<EOT
<div class="mytable w-100 bg-white p-3">
	<div class="myrow">
		<div class="mycell" style="width:20%;">
		</div>
		<div class="mycell weight pt-5 text-center">
			<h3 class="pt-3">團隊支援查詢</h3>
		</div>
		<div class="mycell text-end p-2 vbottom" style="width:20%;">
			<div class="btn-group print"  role="group" style="position:fixed;top: 10px; right:10px;z-index: 9999;">
				<button id="close" class="btn btn-info btn-lg" type="button" onclick="window.print();"><i class="bi bi-printer"></i>&nbsp;列印</button>
				<button id="close" class="btn btn-danger btn-lg" type="button" onclick="window.close();"><i class="bi bi-power"></i>&nbsp;關閉</button>
			</div>
		</div>
	</div>
</div>
<hr class="style_a m-2 p-0">
<div class="w-100 p-3 m-auto text-center">
	<div class="inline size12 weight text-nowrap vtop mb-2 me-2">團隊 : $select_team</div>
	<div class="inline size12 weight text-nowrap vtop mb-2 me-2">支援團隊 : $select_team2</div>
	<div class="inline size12 weight text-nowrap vtop mb-2 me-2">團隊工地 : $select_team_construction</div>
	<div class="inline mb-2">
		<div class="inline size12 weight text-nowrap pt-2 vtop mb-2">請選擇日期範圍：</div>
		<div class="inline text-nowrap mb-2">
			<div class="input-group" id="startdate" style="width:100%;max-width:180px;">
				<input type="text" class="form-control" id="start_date" name="start_date" placeholder="請輸入起始日期" aria-describedby="start_date" value="$start_date">
				<button class="btn btn-outline-secondary input-group-append input-group-addon" type="button" data-target="#startdate" data-toggle="datetimepicker"><i class="bi bi-calendar"></i></button>
			</div>
			<script type="text/javascript">
				$(function () {
					$('#startdate').datetimepicker({
						locale: 'zh-tw'
						,format:"YYYY-MM-DD"
						,allowInputToggle: true
					});
				});
			</script>
		</div>
		<div class="inline text-nowrap mb-2">
			<div class="input-group" id="enddate" style="width:100%;max-width:180px;">
				<input type="text" class="form-control" id="end_date" name="end_date" placeholder="請輸入迄止日期" aria-describedby="end_date" value="$end_date">
				<button class="btn btn-outline-secondary input-group-append input-group-addon" type="button" data-target="#enddate" data-toggle="datetimepicker"><i class="bi bi-calendar"></i></button>
			</div>
			<script type="text/javascript">
				$(function () {
					$('#enddate').datetimepicker({
						locale: 'zh-tw'
						,format:"YYYY-MM-DD"
						,allowInputToggle: true
					});
				});
			</script>
		</div>
		<div class="inline text-nowrap vtop mb-2 me-2">
			<div class="inline text-nowrap me-2">
				<button type="button" class="btn btn-success" onclick="chdatetime();"><i class="fas fa-check"></i>&nbsp;查詢</button>
			</div>
			<div class="inline text-nowrap">
				<a $show_disabled type="button" class="btn btn-primary" href="/index.php?ch=team_support_inquiry_exportexcel&team_id=$team_id&team_id2=$team_id2&team_construction_id=$team_construction_id&start_date=$start_date&end_date=$end_date&fm=$fm"><i class="bi bi-filetype-xls"></i>&nbsp;匯出Excel</a>
			</div>
		</div>
	</div>
</div>
<div class="w-100 px-3 mb-5">
	<div class="text-center">
		<div class="w-100 text-end size12 weight pe-2">單位：工時(日)</div>
		$show_inquiry
	</div>
</div>
EOT;

$show_center=<<<EOT

$show_report

<script>

	function chdatetime() {
		var start_date = $('#start_date').val();
		var end_date = $('#end_date').val();

		var team_id = $('#team_list').val();
		var team_id2 = $('#team2_list').val();
		var team_construction_id = $('#team_construction_list').val();

		window.location = '/index.php?ch=$ch&team_id='+team_id+'&team_id2='+team_id2+'&team_construction_id='+team_construction_id+'&start_date='+start_date+'&end_date='+end_date+'&fm=$fm';
		return false;

	}	

</script>
EOT;



?>