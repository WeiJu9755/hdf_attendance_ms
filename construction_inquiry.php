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
//$project_id = $_GET['project_id'];
//$auth_id = $_GET['auth_id'];

$company_id = $_GET['company_id'];
$construction_id = $_GET['construction_id'];
$get_construction_id = $_GET['construction_id'];
$attendance_status = $_GET['attendance_status'];


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

//載入公司
if  ((($super_advanced=="Y") && ($advanced_readonly <> "Y")) && ($super_admin <> "Y")) {
	$Qry="SELECT a.company_id,a.company_name FROM company a
	RIGHT JOIN group_company b ON b.company_id = a.company_id and b.member_no = '$memberID'
	WHERE a.company_id <> ''
	ORDER BY a.company_id";
} else {
	$Qry="SELECT company_id,company_name FROM company ORDER BY company_id";
}

$mDB->query($Qry);


$select_company = "";
$select_company  = "<select class=\"inline form-select\" name=\"company_list\" id=\"company_list\" style=\"width:auto;\">";
$select_company .= "<option></option>";

if ($mDB->rowCount() > 0) {
	while ($row=$mDB->fetchRow(2)) {
		$ch_company_id = $row['company_id'];
		$ch_company_name = $row['company_name'];
		$select_company .= "<option value='$ch_company_id' ".mySelect($ch_company_id,$company_id).">$ch_company_name $ch_company_id</option>";
	}
}
$select_company .= "</select>";

// 載入工地

$Qry="SELECT construction_id,construction_site from construction";

$mDB->query($Qry);


$select_construction  = "<select class=\"inline form-select\" name=\"construction_id_list\" id=\"construction_id_list\" style=\"width:auto;\">";
$select_construction .= "<option></option>";

if ($mDB->rowCount() > 0) {
	while ($row=$mDB->fetchRow(2)) {
		$ch_construction_id = $row['construction_id'];
		$ch_construction_site = $row['construction_site'];
		$select_construction .= "<option value='$ch_construction_id' ".mySelect($ch_construction_id,$construction_id).">$ch_construction_site $ch_construction_id</option>";
	}
}
$select_construction .= "</select>";

// 載入類別
$selected_status = $_GET['attendance_status'] ?? ''; 

$Qry = "SELECT DISTINCT attendance_status FROM dispatch_construction;";
$mDB->query($Qry);

$select_attendance_status  = "<select class=\"inline form-select\" name=\"attendance_status\" id=\"attendance_status\" style=\"width:auto;\">";
$select_attendance_status .= "<option value='' ".($selected_status == '' ? 'selected' : '')."></option>";

if ($mDB->rowCount() > 0) {
	while ($row = $mDB->fetchRow(2)) {
		$attendance_status = $row['attendance_status'];
		$selected = ($attendance_status == $selected_status) ? 'selected' : '';
		$select_attendance_status .= "<option value='$attendance_status' $selected>$attendance_status</option>";
	}
}
$select_attendance_status .= "</select>";


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
    $m_header[] = [$date->format('Y-m-d'),$date->format('d')];
    $m_manpower_total[] = 0;
    $m_attendance_day_total[] = 0;
	$m_manpower[] = 0;
	$m_attendance_day[] = 0;
}

// 輸出日期陣列
//print_r($dateArray);
//exit;


$show_disabled = "style=\"pointer-events: none;\"";

$show_inquiry = "";

//取得工地資料
if($_GET['company_id']!=""){
$Qry="SELECT a.dispatch_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day
,b.construction_id,c.construction_site,b.building,b.household,b.floor,b.attendance_status,b.team_id,d.team_name,b.manpower,b.workinghours FROM dispatch a
LEFT JOIN dispatch_construction b ON b.dispatch_id = a.dispatch_id
LEFT JOIN construction c ON c.construction_id = b.construction_id
LEFT JOIN team d ON d.team_id = b.team_id
WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' AND a.company_id = '$company_id'";



$Qry .="GROUP BY b.construction_id
		ORDER BY b.construction_id";

}else{
	$Qry="SELECT a.dispatch_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day
,b.construction_id,c.construction_site,b.building,b.household,b.floor,b.attendance_status,b.team_id,d.team_name,b.manpower,b.workinghours FROM dispatch a
LEFT JOIN dispatch_construction b ON b.dispatch_id = a.dispatch_id
LEFT JOIN construction c ON c.construction_id = b.construction_id
LEFT JOIN team d ON d.team_id = b.team_id
WHERE b.construction_id = '{$_GET['construction_id']}'";



$Qry .="GROUP BY b.construction_id
		ORDER BY b.construction_id";
}
//echo $Qry;
//exit; 

$mDB->query($Qry);


$manpower_TOTAL = 0;
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
				<th scope="col" class="size12 bg-aqua text-nowrap" style="padding: 10px;width:2%;background-color: #D2F2FF;"><b>序號</b></th>
				<th scope="col" class="size12 bg-aqua text-nowrap" style="padding: 10px;width:5%;background-color: #D2F2FF;"><b>工地</b></th>
EOT;

$i_count = count($m_header);
for ($i = 0; $i < $i_count; $i++) {
    $show_inquiry.="<th scope=\"col\" class=\"size14 bg-aqua\" style=\"width:3%;padding: 10px;background-color: #D2F2FF;\"><b>".$m_header[$i][1]."</b></th>";
}

$show_inquiry.=<<<EOT
				<th scope="col" class="size14 bg-aqua" style="padding: 10px;width:5%;background-color: #D2F2FF;"><b>合計</b></th>
			</tr>
		</thead>
		<tbody>
EOT;

	$seq = 0;
	while ($row=$mDB->fetchRow(2)) {
		
		$dispatch_id = $row['dispatch_id'];
		$dispatch_year = $row['dispatch_year'];
		$dispatch_month = $row['dispatch_month'];
		$dispatch_day = $row['dispatch_day'];
		$construction_id = $row['construction_id'];
		$construction_site = $row['construction_site'];
		

		$seq++;

		//再取得各員工的資料
		$Qry2="SELECT a.dispatch_id,a.dispatch_date,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day
			,b.construction_id,c.construction_site,b.building,b.household,b.floor,b.attendance_status,b.team_id,d.team_name,b.manpower,b.workinghours
			FROM dispatch a
			LEFT JOIN dispatch_construction b ON b.dispatch_id = a.dispatch_id
			LEFT JOIN construction c ON c.construction_id = b.construction_id
			LEFT JOIN team d ON d.team_id = b.team_id
			WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' AND b.construction_id = '$construction_id'";

		if(!empty($selected_status)){
			$Qry2 .= " AND b.attendance_status = '$selected_status'";
		}	
	
	
		$Qry2 .= "	ORDER BY b.construction_id";

		//echo $Qry2;
		//exit;

		$mDB2->query($Qry2);
		$manpower_sum = 0;
		$attendance_day_sum = 0;
		if ($mDB2->rowCount() > 0) {

			$m_manpower = [];
			foreach ($period as $date) {
				$m_manpower[] = "&nbsp;";
			}

			$m_attendance_day = [];
			foreach ($period as $date) {
				$m_attendance_day[] = "";
			}


			while ($row2=$mDB2->fetchRow(2)) {

				$DAY = $row2['dispatch_date'];

				$i_count = count($m_attendance_day);
				for ($i = 0; $i < $i_count; $i++) {

					if ($DAY == $m_header[$i][0]) {

						//$status_str = "";

						//$status_str .= "<span class=\"text-nowrap\">".$row2['work_overtime'].$work_overtime_hours."</span>";
						$manpower = $row2['manpower'];
						$workinghours = round($row2['workinghours']/8,4);

						$m_attendance_day[$i] .= "<div class=\"text-nowrap\">".$row2['attendance_status'].$row2['team_name']."</div>"
								."<div class=\"text-nowrap mb-2\">"."(<i class=\"bi bi-person-raised-hand\" title=\"人力\"></i>:".$row2['manpower'].")"
								."(<i class=\"bi bi-clock\" title=\"工時/日\"></i>:".$workinghours.")"."</div>";
						

						$manpower_sum = $manpower_sum+$manpower;					//橫向加總
						$attendance_day_sum = $attendance_day_sum+$workinghours;	//橫向加總

						//累計加總
						$m_manpower_total[$i] = $m_manpower_total[$i]+$manpower;
						$m_attendance_day_total[$i] = $m_attendance_day_total[$i]+$workinghours;
						
					}

				}

			}
		}

		$manpower_TOTAL = $attendance_day_TOTAL+$manpower_sum;
		$attendance_day_TOTAL = $attendance_day_TOTAL+$attendance_day_sum;


$show_inquiry.=<<<EOT
			<tr class="text-center">
				<td scope="row" class="text-nowrap vtop">$seq</td>
				<td scope="row" class="text-nowrap vtop"><div class="size14 weight">$construction_site</div><div class="size08">$construction_id</div></td>
EOT;

$i_count = count($m_attendance_day);
for ($i = 0; $i < $i_count; $i++) {
    $show_inquiry.="<td class=\"text-center vtop\">".$m_attendance_day[$i]."</td>";
}

//<td rowspan="2" class="text-end size12 weight vmiddle"><i>$fmt_SUMMARY</i></td>

$show_inquiry.=<<<EOT
				<td scope="row" class="text-nowrap vtop"><div class="weight">總人數：{$manpower_sum}</div><div class="weight">總工數：{$attendance_day_sum}</div></td>
			</tr>
EOT;		
/*
$i_count = count($m_attendance_day);
for ($i = 0; $i < $i_count; $i++) {
    $show_inquiry.="<td class=\"text-center vmiddle\" style=\"width:70px;\">".$m_manpower[$i]."</td>";
}

$show_inquiry.=<<<EOT
			<tr>
EOT;		
*/
		
	}


//顯示全部總和
$show_inquiry.=<<<EOT
   <tr class="text-center bg-yellow size14 weight" style="border-top: 2px solid #000;">
	   <td scope="row" rowspan="2" class="text-nowrap text-center vmiddle" style="background-color: #FFEBAC;">合計</td>
	   <td scope="row" class="text-nowrap text-center vmiddle" style="background-color: #FFEBAC;">總人數</td>
EOT;

$i_count = count($m_manpower_total);
for ($i = 0; $i < $i_count; $i++) {
    $show_inquiry.="<td class=\"text-nowrap text-center vmiddle\" style=\"width:70px;padding: 10px 0;background-color: #FFEBAC;\">".$m_manpower_total[$i]."</td>";
}

$show_inquiry.=<<<EOT
	   <td class="text-end size12 weight text-nowrap vmiddle" style="background-color: #FFEBAC;"><i>$manpower_TOTAL</i></td>
   </tr>
EOT;			

$show_inquiry.=<<<EOT
   <tr class="text-center bg-yellow size14 weight" style="border-top: 2px solid #000;">
	   <td scope="row" class="text-nowrap text-center vmiddle" style="background-color: #FFEBAC;">總工數</td>
EOT;

$i_count = count($m_attendance_day_total);
for ($i = 0; $i < $i_count; $i++) {
    $show_inquiry.="<td class=\"text-nowrap text-center vmiddle\" style=\"width:70px;padding: 10px 0;background-color: #FFEBAC;\">".$m_attendance_day_total[$i]."</td>";
}

$show_inquiry.=<<<EOT
	   <td class="text-end size12 weight text-nowrap vmiddle" style="background-color: #FFEBAC;"><i>$attendance_day_TOTAL</i></td>
   </tr>
EOT;			

$show_inquiry.=<<<EOT
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
			<h3>工地狀況查詢</h3>
		</div>
		<div class="mycell text-end p-2 vbottom" style="width:20%;">
			<div class="w-auto" style="position:fixed;top: 10px; right:10px;z-index: 9999;">
			</div>
		</div>
	</div>
</div>
<hr class="style_a m-2 p-0">
<div class="w-100 p-3 m-auto text-center">
	<div class="inline size12 weight text-nowrap vtop mb-2 me-2">公司 : $select_company</div>
	<div class="inline size12 weight text-nowrap vtop mb-2 me-2">工地 : $select_construction</div>
	<div class="inline size12 weight text-nowrap vtop mb-2 me-2">類別 : $select_attendance_status</div>
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
				<a $show_disabled type="button" class="btn btn-primary" href="/index.php?ch=construction_inquiry_exportexcel&company_id=$company_id&construction_id=$get_construction_id &attendance_status=$selected_status&start_date=$start_date&end_date=$end_date&fm=$fm"><i class="bi bi-filetype-xls"></i>&nbsp;匯出Excel</a>
			</div>
		</div>
	</div>
</div>
<div class="w-100 px-3 mb-5">
	<div class="overflow-auto" style="white-space: nowrap;">
		<div class="text-center">
			<div class="w-100 text-end size12 weight pe-2">單位：工時(日)</div>
			$show_inquiry
		</div>
	</div>
</div>
EOT;

$show_center=<<<EOT

$show_report

<script>

	function chdatetime() {
		var start_date = $('#start_date').val();
		var end_date = $('#end_date').val();
		var company_id = $('#company_list').val();
		var construction_id = $('#construction_id_list').val();
		var attendance_status = $('#attendance_status').val();

		window.location = '/index.php?ch=$ch&company_id=' + company_id +
						  '&construction_id=' + construction_id +
						  '&attendance_status=' + attendance_status +
						  '&start_date=' + start_date +
						  '&end_date=' + end_date +
						  '&fm=$fm';
		return false;

		/*
		// 取得今天的日期
        const today = new Date();
        today.setHours(23, 59, 59, 59); // 清除時間部分，以確保只有日期被比較

		const inputDate = new Date(end_date);
		// 檢查輸入日期是否超過今天
        if (inputDate > today) {
            event.preventDefault(); // 阻止表單提交
			jAlert('警示', '輸入日期不能超過今天', 'red', '', 2000);
			return false;
        } else {
			//alert(start_date+' '+end_date);
			window.location = '/index.php?ch=$ch&construction_id=$construction_id&start_date='+start_date+'&end_date='+end_date+'&fm=$fm';
			return false;
		}
			*/
	}	

</script>
EOT;



?>