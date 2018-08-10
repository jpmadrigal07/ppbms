<?php
include_once("../includes/loginstatus.php");

$rpp = preg_replace('#[^0-9]#', '', $_POST['rpp']);
$last = preg_replace('#[^0-9]#', '', $_POST['last']);
$pn = preg_replace('#[^0-9]#', '', $_POST['pn']);
// This makes sure the page number isn't below 1, or more than our $last page
if ($pn < 1) { 
    $pn = 1; 
} else if ($pn > $last) { 
    $pn = $last; 
}
// This sets the range of rows to query for the chosen $pn
$limit = 'LIMIT ' .($pn - 1) * $rpp .',' .$rpp;

$output = '';
$count = 0; 

    $sql = "SELECT id, encode_file_name, encode_lists_date, encode_lists_date FROM ppbms_encode_list WHERE encode_lists_status = '1' ORDER BY id DESC $limit";  
  $result = mysqli_query($db_conn, $sql);
  
  while($row = mysqli_fetch_array($result, MYSQLI_ASSOC)) {  
    $count++; 
    $recid = $row["id"] == "" ? 'empty' : $row["id"];
    $filename = $row["encode_file_name"] == "" ? 'empty' : $row["encode_file_name"];
    $date = $row["encode_lists_date"] == "" ? 'empty' : $row["encode_lists_date"];

    $newDate = date("M d, Y - g:i A", strtotime($date));
    $newDateUpdate = date("m/d/Y H:i A", strtotime($date));

    $sql_unassined = "SELECT COUNT(id) AS count_num FROM ppbms_record WHERE encode_lists_id = '$recid' AND messenger_id = '0' AND record_status_status = '1'";
    $query_unassined = mysqli_query($db_conn, $sql_unassined);
    while($row1 = mysqli_fetch_array($query_unassined)) {  
      $unassined_count = $row1["count_num"];
    }

    $sql_assined = "SELECT COUNT(id) AS count_num FROM ppbms_record WHERE encode_lists_id = '$recid' AND messenger_id != '0' AND record_status_status = '1'";
    $query_assined = mysqli_query($db_conn, $sql_assined);
    while($row2 = mysqli_fetch_array($query_assined)) {  
      $assined_count = $row2["count_num"];
    }

    $sql_total = "SELECT COUNT(id) AS count_num FROM ppbms_record WHERE encode_lists_id = '$recid' AND record_status_status = '1'";
    $query_total= mysqli_query($db_conn, $sql_total);
    while($row3 = mysqli_fetch_array($query_total)) {  
      $total_count = $row3["count_num"];
    }
                      
    $output .= $recid.'|'.$filename.'|'.$newDate.'|'.$assined_count.'|'.$unassined_count.'|'.$total_count.'||';
  }

  echo $output;

  mysqli_close($db_conn);

  exit();   

?>