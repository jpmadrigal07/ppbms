<?php
header('Content-Type: application/json');
include_once("../includes/loginstatus.php");
	// This first query is just to get the total count of rows

	$sql = "SELECT COUNT(id) AS count_num FROM ppbms_dispatch_control_messenger WHERE dispatch_control_messenger_status = '1'";  
	$result = mysqli_query($db_conn, $sql);
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