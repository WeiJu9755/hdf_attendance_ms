<?php

		$m_location		= "/website/smarty/templates/".$site_db."/".$templates;
		$m_pub_modal	= "/website/smarty/templates/".$site_db."/pub_modal";

		//程式分類
		$ch = empty($_GET['ch']) ? 'default' : $_GET['ch'];
		switch($ch) {
			case 'attendance_inquiry':
				$title = "出勤查詢";
				$sid = "view01";
				$modal = $m_location."/sub_modal/project/func08/attendance_ms/attendance_inquiry.php";
				include $modal;
				$smarty->assign('show_center',$show_center);
				break;
			case 'attendance_inquiry_exportexcel':
				$title = "匯出Excel";
				$sid = "view01";
				$modal = $m_location."/sub_modal/project/func08/attendance_ms/attendance_inquiry_exportexcel.php";
				include $modal;
				$smarty->assign('show_center',$show_center);
				break;
			case 'team_support_inquiry':
				$title = "團隊支援查詢";
				$sid = "view01";
				$modal = $m_location."/sub_modal/project/func08/attendance_ms/team_support_inquiry.php";
				include $modal;
				$smarty->assign('show_center',$show_center);
				break;
			case 'team_support_inquiry_exportexcel':
				$title = "匯出Excel";
				$sid = "view01";
				$modal = $m_location."/sub_modal/project/func08/attendance_ms/team_support_inquiry_exportexcel.php";
				include $modal;
				$smarty->assign('show_center',$show_center);
				break;
			case 'construction_inquiry':
				$title = "工地狀況查詢";
				$sid = "view01";
				$modal = $m_location."/sub_modal/project/func08/attendance_ms/construction_inquiry.php";
				include $modal;
				$smarty->assign('show_center',$show_center);
				break;
			case 'construction_inquiry_exportexcel':
				$title = "匯出Excel";
				$sid = "view01";
				$modal = $m_location."/sub_modal/project/func08/attendance_ms/construction_inquiry_exportexcel.php";
				include $modal;
				$smarty->assign('show_center',$show_center);
				break;
			default:
				$sid = "mbpjitem";
				$modal = $m_location."/sub_modal/project/func08/attendance_ms/dispatchreport.php";
				include $modal;
				$smarty->assign('show_center',$show_center);
				break;
		};

?>