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

    $sql = "SELECT * FROM ppbms_record WHERE messenger_id = '0' AND record_status_status = '1' AND encode_lists_id = '".$_POST["encodeid"]."' ORDER BY id $limit";  
  $result = mysqli_query($db_conn, $sql);
  
  while($row = mysqli_fetch_array($result, MYSQLI_ASSOC)) {  
    $count++; 
    $recid = $row["id"] == "" ? 'empty' : $row["id"];
      $sender = $row["record_sender"] == "" ? 'empty' : $row["record_sender"];
      $deltype = $row["record_deltype"] == "" ? 'empty' : $row["record_deltype"];
      $pud = $row["record_pud"] == "" ? 'empty' : $row["record_pud"];
      $month = $row["record_month"] == "" ? 'empty' : $row["record_month"];
      $jobnum = $row["record_job_num"] == "" ? 'empty' : $row["record_job_num"];
      $checklist = $row["record_check_list"] == "" ? 'empty' : $row["record_check_list"];
      $filename = $row["record_file_name"] == "" ? 'empty' : $row["record_file_name"];
      $seqnum = $row["record_seq_num"] == "" ? 'empty' : $row["record_seq_num"];
      $cyclecode = $row["record_cycle_code"] == "" ? 'empty' : $row["record_cycle_code"];
      $qty = $row["record_qty"] == "" ? 'empty' : $row["record_qty"];
      $address = $row["record_address"] == "" ? 'empty' : $row["record_address"];
      $area = $row["record_area"] == "" ? 'empty' : $row["record_area"];
      $subs = $row["record_subs_name"] == "" ? 'empty' : $row["record_subs_name"];
      $barcode = $row["record_bar_code"] == "" ? 'empty' : $row["record_bar_code"];
      $acctnum = $row["record_acct_num"] == "" ? 'empty' : $row["record_acct_num"];
      $datereceive = $row["record_date_receive"] == "" ? 'empty' : $row["record_date_receive"];
      $receiveby = $row["record_receive_by"] == "" ? 'empty' : $row["record_receive_by"];
      $relation = $row["record_relation"] == "" ? 'empty' : $row["record_relation"];
      $messenger = $row["record_messenger"] == "" ? 'empty' : $row["record_messenger"];
      $status = $row["record_status"] == "" ? 'empty' : $row["record_status"];
      $reasonrts = $row["record_reason_rts"] == "" ? 'empty' : $row["record_reason_rts"];
      $remarks = $row["record_remarks"] == "" ? 'empty' : $row["record_remarks"];
      $datereported = $row["record_date_reported"] == "" ? 'empty' : $row["record_date_reported"];

    $output .= $recid.'|'.$sender.'|'.$deltype.'|'.$pud.'|'.$month.'|'.$jobnum.'|'.$checklist.'|'.$filename.'|'.$seqnum.'|'.$cyclecode.'|'.$qty.'|'.$address.'|'.$area.'|'.$subs.'|'.$barcode.'|'.$acctnum.'|'.$datereceive.'|'.$receiveby.'|'.$relation.'|'.$messenger.'|'.$status.'|'.$reasonrts.'|'.$remarks.'|'.$datereported.'||';
  }

  echo $output;

  mysqli_close($db_conn);

  exit();   

?>