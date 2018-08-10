<?php  
include_once("../includes/loginstatus.php");

$output = '';
$count = 0; 

$sender = $_POST["sender"];
$deltype = $_POST["deltype"];
$month = $_POST["month"];
$year = $_POST["year"];

$sql = "SELECT DISTINCT(record_cycle_code), record_pud FROM ppbms_record WHERE record_status_status = '1' AND record_year = '$year' AND  record_month = '$month' AND record_sender = '$sender' AND record_deltype = '$deltype' ORDER BY id DESC";  
$result = mysqli_query($db_conn, $sql);  

  $output .= '
  <form method="post" action="reportdays.php">
  <input name="sender" type="hidden" value="'.$sender.'" />
  <input name="deltype" type="hidden" value="'.$deltype.'" />
  <input name="month" type="hidden" value="'.$month.'" />
  <input name="year" type="hidden" value="'.$year.'" />
  <table class="table table-striped">
  <thead>
  <tr>
  <th></th>
  <th>Cycle Code</th>
  <th>Pick-up Date</th>
  </tr>
  </thead>
  <tbody>
  ';
  while($row = mysqli_fetch_array($result)) {  
    $count++; 
    $cyclecode = $row["record_cycle_code"];
    $pud = $row["record_pud"];


    $output .= '
    <tr>
    <td><input name="checkboxsummary[]" type="checkbox" value="'.$cyclecode.','.$pud.'"></td>        
    <td>'.$cyclecode.'</td>
    <td>'.$pud.'</td>
    </tr>
    ';  
  }
  $output .= '
  </tbody>
  </table>
  <input type="submit" class="btn btn-primary" style="width: 100%;" value="Export">
  </form>
  ';
  echo $output; 

  mysqli_close($db_conn);

  exit(); 

?>  