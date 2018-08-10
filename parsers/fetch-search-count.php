<?php
header('Content-Type: application/json');
include_once("../includes/loginstatus.php");
	// This first query is just to get the total count of rows

	if($_POST["filter"] == "all") {  
	  $sql = "SELECT COUNT(id) AS count_num FROM ppbms_record WHERE record_status_status = '1' AND encode_lists_id = '".$_POST["encodeid"]."' AND (record_subs_name LIKE '%".$_POST["search"]."%' OR record_cycle_code LIKE '%".$_POST["search"]."%' OR record_bar_code LIKE '%".$_POST["search"]."%' OR record_deltype LIKE '%".$_POST["search"]."%' OR record_sender LIKE '%".$_POST["search"]."%')";  
	  $result = mysqli_query($db_conn, $sql);  
	} else if($_POST["filter"] == "cyclecode") {  
	  $sql = "SELECT COUNT(id) AS count_num FROM ppbms_record WHERE record_status_status = '1' AND encode_lists_id = '".$_POST["encodeid"]."' AND record_cycle_code LIKE '%".$_POST["search"]."%'";  
	  $result = mysqli_query($db_conn, $sql);  
	} else if($_POST["filter"] == "barcode") {  
	  $sql = "SELECT COUNT(id) AS count_num FROM ppbms_record WHERE record_status_status = '1' AND encode_lists_id = '".$_POST["encodeid"]."' AND record_bar_code LIKE '%".$_POST["search"]."%'";  
	  $result = mysqli_query($db_conn, $sql);  
	} else if($_POST["filter"] == "deltype") {  
	  $sql = "SELECT COUNT(id) AS count_num FROM ppbms_record WHERE record_status_status = '1' AND encode_lists_id = '".$_POST["encodeid"]."' AND record_deltype LIKE '%".$_POST["search"]."%'";  
	  $result = mysqli_query($db_conn, $sql);
	} else if($_POST["filter"] == "sender") {  
	  $sql = "SELECT COUNT(id) AS count_num FROM ppbms_record WHERE record_status_status = '1' AND encode_lists_id = '".$_POST["encodeid"]."' AND record_sender LIKE '%".$_POST["search"]."%'";  
	  $result = mysqli_query($db_conn, $sql);  
	} else if($_POST["filter"] == "name") {  
	  $sql = "SELECT COUNT(id) AS count_num FROM ppbms_record WHERE record_status_status = '1' AND encode_lists_id = '".$_POST["encodeid"]."' AND record_subs_name LIKE '%".$_POST["search"]."%'";  
	  $result = mysqli_query($db_conn, $sql);  
	}

	while($row = mysqli_fetch_array($result)) {  
		$total = $row["count_num"];
	}
	// Specify how many results per page
	$rpp = 10;
	// This tells us the page number of our last page
	$last = ceil($total/$rpp);
	// This makes sure $last cannot be less than 1
	if($last < 1){
		$last = 1;
	}

	echo $last;

  mysqli_close($db_conn);

  exit(); 

	
?>