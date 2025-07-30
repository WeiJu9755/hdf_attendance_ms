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
$team_id = $_GET['team_id'];


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


$getteamclass = "/smarty/templates/$site_db/$templates/sub_modal/project/func08/attendance_ms/getteamclass.php";


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



//載入團隊
$Qry="SELECT team_id,team_name FROM team WHERE company_id = '$company_id' ORDER BY team_id";
$select_team = "";
$select_team  = "<select class=\"inline form-select\" name=\"team_list\" id=\"team_list\" style=\"width:auto;min-width: 150px;\">";
$select_team .= "<option></option>";
$mDB->query($Qry);
if ($mDB->rowCount() > 0) {
	while ($row=$mDB->fetchRow(2)) {
		$ch_team_id = $row['team_id'];
		$ch_team_name = $row['team_name'];
		$select_team .= "<option value=\"$ch_team_id\" ".mySelect($ch_team_id,$team_id).">$ch_team_name $ch_team_id</option>";
	}
}	
$select_team .= "</select>";





//計算日期間隔
//$date1 = new DateTime($start_date);
//$date2 = new DateTime($end_date);
//$day_interval = $date1->diff($date2);


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

// 輸出日期陣列
//print_r($dateArray);
//exit;


$show_disabled = "style=\"pointer-events: none;\"";

$show_inquiry = "";


//取得工單資料
if (!empty($team_id)) {
	$Qry="SELECT a.dispatch_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day,b.employee_id,c.employee_name,c.team_id,d.team_name FROM dispatch a
	LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
	LEFT JOIN employee c ON c.employee_id = b.employee_id
	LEFT JOIN team d ON d.team_id = c.team_id
	WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' AND a.company_id = '$company_id' AND c.team_id = '$team_id'
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

//echo $Qry;
//exit; 

$mDB->query($Qry);


$attendance_day_TOTAL = 0;

$alist_kwh = array();


if ($mDB->rowCount() > 0) {

	$show_disabled = "";

	//顯示抬頭標題列
$show_inquiry.=<<<EOT
	<table class="table table-bordered" style="border: 2px solid #000;background-color: #FFFFFF;">
		<thead>
			<tr class="text-center" style="border-bottom: 2px solid #000;">
				<th scope="col" class="size12 bg-aqua text-nowrap" style="padding: 10px;width:2%;background-color: #D2F2FF;"><b>序號</b></th>
				<th scope="col" class="size12 bg-aqua text-nowrap" style="padding: 10px;width:5%;background-color: #D2F2FF;"><b>團隊</b></th>
				<th scope="col" class="size12 bg-aqua text-nowrap" style="padding: 10px;width:5%;background-color: #D2F2FF;"><b>員工</b></th>
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
		$employee_id = $row['employee_id'];
		$employee_name = $row['employee_name'];
		$m_team_id = $row['team_id'];
		$m_team_name = $row['team_name'];

		if (!empty($employee_id)) {
			$seq++;

		//再取得各員工的資料
		$Qry2="SELECT a.dispatch_id,a.dispatch_date,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day
			,b.employee_id,b.manpower,b.attendance_day,b.attendance_status,b.leave,b.leave_start,b.leave_end,b.work_overtime,b.work_overtime_start,b.work_overtime_end
			,b.attendance_end_status,b.attendance_end,b.transition,b.transition_team_id,b.transition_construction_id,b.transition_start,b.transition_end
			FROM dispatch a
			LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
			WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' AND b.employee_id = '$employee_id'
			ORDER BY b.employee_id";

		//echo $Qry2;
		//exit;

		$mDB2->query($Qry2);
		$SUMMARY = 0;
		if ($mDB2->rowCount() > 0) {

			$m_attendance_day = [];
			foreach ($period as $date) {
				$m_attendance_day[] = "&nbsp;";
			}

			$m_leave = [];
			foreach ($period as $date) {
				$m_leave[] = "&nbsp;";
			}


			while ($row2=$mDB2->fetchRow(2)) {

				$DAY = $row2['dispatch_date'];

				$i_count = count($m_attendance_day);
				for ($i = 0; $i < $i_count; $i++) {

					if ($DAY == $m_header[$i][0]) {

						$status_str = "";

						if ($row2['attendance_status'] == "颱風假") {
							$status_str .= "<span class=\"text-nowrap\">".$row2['attendance_status']."</span>";
						}

						if ($row2['transition'] == "Y") {
							//計算轉場時數/日
							$transition_start = $row2['transition_start'];
							$transition_end = $row2['transition_end'];

							$transition_hours = round(calculateLeaveHours($transition_start, $transition_end)/8,4);

							$attendance_end = substr($row2['attendance_end'],0,5);
							$status_str .= "<span class=\"text-nowrap\">轉場".$transition_hours."</span>";
						}

						if (!empty($row2['leave'])) {
							//計算請假時數/日
							$leave_start = $row2['leave_start'];
							$leave_end = $row2['leave_end'];

							$leave_hours = round(calculateLeaveHours($leave_start, $leave_end)/8,4);

							$status_str .= "<span class=\"text-nowrap\">".$row2['leave'].$leave_hours."</span>";
						}

						if ($row2['attendance_end_status'] == "提早收工") {
							$attendance_end = substr($row2['attendance_end'],0,5);
							$status_str .= "<span class=\"text-nowrap\">".$attendance_end.$row2['attendance_end_status']."</span>";
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

							$status_str .= "<span class=\"text-nowrap\">".$row2['work_overtime'].$work_overtime_hours."</span>";
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

		$attendance_day_TOTAL = $attendance_day_TOTAL+$SUMMARY;

		$i_count = count($m_attendance_day_TOT);
		for ($i = 0; $i < $i_count; $i++) {
			if ($m_attendance_day_TOT[$i] == 0) {
				$m_attendance_day_TOT[$i] = number_format2($m_attendance_day_TOT[$i],4);
			}
		}

		$fmt_SUMMARY = number_format2($SUMMARY,4);

		$fmt_attendance_day_TOTAL = number_format2($attendance_day_TOTAL,4);


$show_inquiry.=<<<EOT
			<tr class="text-center">
				<td scope="row" rowspan="2" class="text-nowrap vmiddle">$seq</td>
				<td scope="row" rowspan="2" class="text-nowrap vmiddle"><div class="size14 weight">$m_team_name</div><div class="size08">$m_team_id</div></td>
				<td scope="row" rowspan="2" class="vmiddle"><div class="size14 weight">$employee_name</div><div class="size08">$employee_id</div></td>
EOT;

$i_count = count($m_attendance_day);
for ($i = 0; $i < $i_count; $i++) {
    $show_inquiry.="<td class=\"text-center vmiddle\">".$m_attendance_day[$i]."</td>";
}


$show_inquiry.=<<<EOT
				<td rowspan="2" class="text-end size12 weight vmiddle"><i>$fmt_SUMMARY</i></td>
			</tr>
			<tr>
EOT;		

$i_count = count($m_attendance_day);
for ($i = 0; $i < $i_count; $i++) {
    $show_inquiry.="<td class=\"text-center vmiddle\">".$m_leave[$i]."</td>";
}

$show_inquiry.=<<<EOT
			<tr>
EOT;		


		}

	}


//顯示全部總和
$show_inquiry.=<<<EOT
   <tr class="text-center bg-yellow size14 weight" style="border-top: 2px solid #000;">
	   <td scope="row" colspan="3" class="text-nowrap text-center vmiddle" style="background-color: #FFEBAC;">合計</td>
EOT;

$i_count = count($m_attendance_day_TOT);
for ($i = 0; $i < $i_count; $i++) {
    $show_inquiry.="<td class=\"text-nowrap text-center vmiddle\" style=\"padding: 10px;background-color: #FFEBAC;\">".$m_attendance_day_TOT[$i]."</td>";
}

$show_inquiry.=<<<EOT
	   <td class="text-end size12 weight text-nowrap vmiddle" style="background-color: #FFEBAC;"><i>$fmt_attendance_day_TOTAL</i></td>
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
<style>

table.table-bordered {
	border:1px solid black;
}
table.table-bordered > thead > tr > th{
	border:1px solid black;
}
table.table-bordered > tbody > tr > th {
	border:1px solid black;
}
table.table-bordered > tbody > tr > td {
	border:1px solid black;
}

@media print {
	.print {
		display: none !important;
	}
}

</style>

<div class="mytable w-100 bg-white p-3">
	<div class="myrow">
		<div class="mycell" style="width:20%;">
		</div>
		<div class="mycell weight pt-5 text-center">
			<h3 class="pt-3">出勤查詢</h3>
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
	<div class="inline size12 weight text-nowrap vtop mb-2 me-2">公司 : $select_company</div>
	<div class="inline size12 weight text-nowrap vtop mb-2 me-2">團隊 : $select_team</div>
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
		<div class="inline text-nowrap vtop mb-2 me-2 print">
			<div class="inline text-nowrap me-2">
				<button type="button" class="btn btn-success" onclick="chdatetime();"><i class="fas fa-check"></i>&nbsp;查詢</button>
			</div>
			<div class="inline text-nowrap">
				<a $show_disabled type="button" class="btn btn-primary" href="/index.php?ch=attendance_inquiry_exportexcel&company_id=$company_id&team_id=$team_id&start_date=$start_date&end_date=$end_date&fm=$fm"><i class="bi bi-filetype-xls"></i>&nbsp;匯出Excel</a>
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

		var company_id = $('#company_list').val();
		var team_id = $('#team_list').val();

		window.location = '/index.php?ch=$ch&company_id='+company_id+'&team_id='+team_id+'&start_date='+start_date+'&end_date='+end_date+'&fm=$fm';
		return false;

	}	


function getSelectVal(){ 
	$("option",team_list).remove(); //清空原有的選項
	var company_id = $("#company_list").val();
    $.getJSON('$getteamclass',{company_id:company_id,site_db:'$site_db'},function(json){ 
        var team_class = $("#team_list"); 
        var option = "<option></option>";
		team_class.append(option);
        $.each(json,function(index,array){ 
			option = "<option value='"+array['team_id']+"'>"+array['team_name']+" "+array['team_id']+"</option>"; 
            team_class.append(option); 
        }); 
    });
}

$(function(){ 
    $("#company_list").change(function(){ 
        getSelectVal(); 
    }); 
});

/*
//更新主類別
function getMainSelectVal(){ 
    $.getJSON("$getmainclass",{site_db:'$site_db'},function(json){ 
        var main_class = $("#company_list"); 
		var last_option = main_class.val();
        $("option",company_list).remove(); //清空原有的選項
        var option = "<option></option>";
		main_class.append(option);
        $.each(json,function(index,array){
			if (array['caption'] == last_option)
				option = "<option value='"+array['caption']+"' selected>"+array['caption']+"</option>"; 
			else
				option = "<option value='"+array['caption']+"'>"+array['caption']+"</option>"; 
            main_class.append(option); 
        }); 
    }); 
}
*/


</script>
EOT;



?>