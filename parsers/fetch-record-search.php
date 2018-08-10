<?php  
include_once("../includes/loginstatus.php");

$output = '';
$count = 0; 

if($_POST["filter"] == "all") {  
  $sql = "SELECT * FROM ppbms_record WHERE record_status_status = '1' AND record_subs_name LIKE '%".$_POST["search"]."%' OR record_cycle_code LIKE '%".$_POST["search"]."%' OR record_deltype LIKE '%".$_POST["search"]."%' OR record_sender LIKE '%".$_POST["search"]."%' ORDER BY id";  
  $result = mysqli_query($db_conn, $sql);  
} else if($_POST["filter"] == "cyclecode") {  
  $sql = "SELECT * FROM ppbms_record WHERE record_status_status = '1' AND record_cycle_code LIKE '%".$_POST["search"]."%' ORDER BY id";  
  $result = mysqli_query($db_conn, $sql);  
} else if($_POST["filter"] == "deltype") {  
  $sql = "SELECT * FROM ppbms_record WHERE record_status_status = '1' AND record_deltype LIKE '%".$_POST["search"]."%' ORDER BY id";  
  $result = mysqli_query($db_conn, $sql);  
} else if($_POST["filter"] == "sender") {  
  $sql = "SELECT * FROM ppbms_record WHERE record_status_status = '1' AND record_sender LIKE '%".$_POST["search"]."%' ORDER BY id";  
  $result = mysqli_query($db_conn, $sql);  
} else if($_POST["filter"] == "name") {  
  $sql = "SELECT * FROM ppbms_record WHERE record_status_status = '1' AND record_subs_name LIKE '%".$_POST["search"]."%' ORDER BY id";  
  $result = mysqli_query($db_conn, $sql);  
}

if($_POST["search"] == "") {
  $output .= '
  <table class="table table-striped">
  <thead>
  <tr>
  <th>REC#</th>
  <th>SENDER</th>
  <th>DELTYPE</th>
  <th>PUD</th>
  <th>Month</th>
  <th>JOB#</th>
  <th>Checklist For PPB</th>
  <th>File Name</th>
  <th>seq no.</th>
  <th>CYCLECODE</th>
  <th>Qty</th>
  <th>ADDRESS</th>
  <th>AREA</th>
  <th>SUBSCRIBERS NAME</th>
  <th>BARCODE</th>
  <th>ACCT NO</th>
  <th>DATE RECEIVED</th>
  <th>RECEIVED BY</th>
  <th>RELATION</th>
  <th>MESSENGER</th>
  <th>STATUS</th>
  <th>Reason for RTS</th>
  <th>Remarks</th>
  <th>Date Reported</th>
  </tr>
  </thead>
  <tbody>
  ';
  while($row = mysqli_fetch_array($result)) {  
    $count++; 
    $recid = $row["id"];
    $sender = $row["record_sender"];
    $deltype = $row["record_deltype"];
    $pud = $row["record_pud"];
    $month = $row["record_month"];
    $jobnum = $row["record_job_num"];
    $checklist = $row["record_check_list"];
    $filename = $row["record_file_name"];
    $seqnum = $row["record_seq_num"];
    $cyclecode = $row["record_cycle_code"];
    $qty = $row["record_qty"];
    $address = $row["record_address"];
    $area = $row["record_area"];
    $subs = $row["record_subs_name"];
    $barcode = $row["record_bar_code"];
    $acctnum = $row["record_acct_num"];
    $datereceive = $row["record_date_receive"];
    $receiveby = $row["record_receive_by"];
    $relation = $row["record_relation"];
    $messenger = $row["record_messenger"];
    $status = $row["record_status"];
    $reasonrts = $row["record_reason_rts"];
    $remarks = $row["record_remarks"];
    $datereported = $row["record_date_reported"];

    $output .= '
    <tr>
    <td>'.$recid.'</td>
    <td>'.$sender.'</td>
    <td>'.$deltype.'</td>
    <td>'.$pud.'</td>
    <td>'.$month.'</td>
    <td>'.$jobnum.'</td>
    <td>'.$checklist.'</td>
    <td>'.$filename.'</td>
    <td>'.$seqnum.'</td>
    <td>'.$cyclecode.'</td>
    <td>'.$qty.'</td>
    <td>'.$address.'</td>
    <td>'.$area.'</td>
    <td>'.$subs.'</td>
    <td>'.$barcode.'</td>
    <td>'.$acctnum.'</td>
    <td>'.$datereceive.'</td>
    <td>'.$receiveby.'</td>
    <td>'.$relation.'</td>
    <td>'.$messenger.'</td>
    <td>'.$status.'</td>
    <td>'.$reasonrts.'</td>
    <td>'.$remarks.'</td>
    <td>'.$datereported.'</td>
    </tr>
    ';   

  }
  $output .= '
  </tbody>
  </table>
  ';
  // echo $output;  
} else {
  if(mysqli_num_rows($result) > 0) {  
    $output .= '
    <table class="table table-striped">
    <thead>
    <tr>
    <th>REC#</th>
    <th>SENDER</th>
    <th>DELTYPE</th>
    <th>PUD</th>
    <th>Month</th>
    <th>JOB#</th>
    <th>Checklist For PPB</th>
    <th>File Name</th>
    <th>seq no.</th>
    <th>CYCLECODE</th>
    <th>Qty</th>
    <th>ADDRESS</th>
    <th>AREA</th>
    <th>SUBSCRIBERS NAME</th>
    <th>BARCODE</th>
    <th>ACCT NO</th>
    <th>DATE RECEIVED</th>
    <th>RECEIVED BY</th>
    <th>RELATION</th>
    <th>MESSENGER</th>
    <th>STATUS</th>
    <th>Reason for RTS</th>
    <th>Remarks</th>
    <th>Date Reported</th>
    </tr>
    </thead>
    <tbody>
    ';
    while($row = mysqli_fetch_array($result)) {  
      $count++; 
      $recid = $row["id"];
      $sender = $row["record_sender"];
      $deltype = $row["record_deltype"];
      $pud = $row["record_pud"];
      $month = $row["record_month"];
      $jobnum = $row["record_job_num"];
      $checklist = $row["record_check_list"];
      $filename = $row["record_file_name"];
      $seqnum = $row["record_seq_num"];
      $cyclecode = $row["record_cycle_code"];
      $qty = $row["record_qty"];
      $address = $row["record_address"];
      $area = $row["record_area"];
      $subs = $row["record_subs_name"];
      $barcode = $row["record_bar_code"];
      $acctnum = $row["record_acct_num"];
      $datereceive = $row["record_date_receive"];
      $receiveby = $row["record_receive_by"];
      $relation = $row["record_relation"];
      $messenger = $row["record_messenger"];
      $status = $row["record_status"];
      $reasonrts = $row["record_reason_rts"];
      $remarks = $row["record_remarks"];
      $datereported = $row["record_date_reported"];

      $output .= '
      <tr>
      <td>'.$recid.'</td>
      <td>'.$sender.'</td>
      <td>'.$deltype.'</td>
      <td>'.$pud.'</td>
      <td>'.$month.'</td>
      <td>'.$jobnum.'</td>
      <td>'.$checklist.'</td>
      <td>'.$filename.'</td>
      <td>'.$seqnum.'</td>
      <td>'.$cyclecode.'</td>
      <td>'.$qty.'</td>
      <td>'.$address.'</td>
      <td>'.$area.'</td>
      <td>'.$subs.'</td>
      <td>'.$barcode.'</td>
      <td>'.$acctnum.'</td>
      <td>'.$datereceive.'</td>
      <td>'.$receiveby.'</td>
      <td>'.$relation.'</td>
      <td>'.$messenger.'</td>
      <td>'.$status.'</td>
      <td>'.$reasonrts.'</td>
      <td>'.$remarks.'</td>
      <td>'.$datereported.'</td>
      </tr>
      ';   

    }
    $output .= '
    </tbody>
    </table>
    ';
    echo $output;  
  } else {  
  echo '<h4 align="center">No results for <b>'.$_POST["search"].'</b></h4>';  
  } 
}

  mysqli_close($db_conn);

  exit(); 
 
?>  