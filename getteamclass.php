<?

$site_db = $_GET['site_db'];
//$pro_id = $_GET['pro_id'];
$company_id = $_GET['company_id'];

include_once("/website/class/".$site_db."_info_class.php");

//載入公用函數
@include_once '/website/include/pub_function.php';

//先取出 caption () 的 main_class 值

//$m_row = getkeyvalue2($site_db."_info","pjclass","pro_id = '$pro_id' and small_class = '0' and caption = '$main_class'","main_class");
//$main_class_seq = $m_row['main_class'];

	
//從資料庫中讀取主類別資料
$mDB = "";
$mDB = new MywebDB();
//$Qry="SELECT caption FROM pjclass where main_class = '$main_class_seq' and small_class <> '0' order by orderby";
$Qry="SELECT team_id,team_name FROM team WHERE company_id = '$company_id' ORDER BY team_id";
$mDB->query($Qry);
if ($mDB->rowCount() > 0) {
	while ($row=$mDB->fetchRow(2)) {
		$select[] = array(
			"team_id"=>$row[team_id]
			,"team_name"=>$row[team_name]
		); 
	}
	echo json_encode($select); 
}	
$mDB->remove();
?>