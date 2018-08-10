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

    $sql = "SELECT id, dispatch_control_messenger_name, dispatch_control_messenger_address, dispatch_control_messenger_prepared, dispatch_control_messenger_date FROM ppbms_dispatch_control_messenger WHERE dispatch_control_messenger_status = '1' ORDER BY id DESC $limit";  
  $result = mysqli_query($db_conn, $sql);
  
  while($row = mysqli_fetch_array($result, MYSQLI_ASSOC)) {  
    $count++; 
    $recid = $row["id"] == "" ? 'empty' : $row["id"];
    $messengername = $row["dispatch_control_messenger_name"] == "" ? 'empty' : $row["dispatch_control_messenger_name"];
    $address = $row["dispatch_control_messenger_address"] == "" ? 'empty' : $row["dispatch_control_messenger_address"];
    $prepared = $row["dispatch_control_messenger_prepared"] == "" ? 'empty' : $row["dispatch_control_messenger_prepared"];
    $date = $row["dispatch_control_messenger_date"] == "" ? 'empty' : $row["dispatch_control_messenger_date"];
    $newDate = date("F d, Y", strtotime($date));
    $newDateUpdate = date("m/d/Y", strtotime($date));
                      
    $output .= $recid.'|'.$messengername.'|'.$address.'|'.$prepared.'|'.$newDate.'|'.$newDateUpdate.'||';
  }

  echo $output;

  mysqli_close($db_conn);

  exit();   

?>