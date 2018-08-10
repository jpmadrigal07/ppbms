<?php  
include_once("../includes/loginstatus.php");

$output = '';
$count = 0; 

$sql = "SELECT * FROM ppbms_dispatch_control_messenger WHERE (dispatch_control_messenger_name LIKE '%".$_POST["search"]."%' OR dispatch_control_messenger_address LIKE '%".$_POST["search"]."%' OR dispatch_control_messenger_prepared LIKE '%".$_POST["search"]."%') AND dispatch_control_messenger_status = '1' ORDER BY id DESC";  
$result = mysqli_query($db_conn, $sql);  

if($_POST["search"] == "") {
  
  $output .= '
  <div id="pagination_controls_dispatch_1"></div>
  <div id="record-search-result-dispatch"></div>
  <div id="pagination_controls_dispatch_2"></div>
  <script type="text/javascript">
  request_page_dispatch(1);
  </script>
  ';
  echo $output;  
} else {
  if(mysqli_num_rows($result) > 0) {  
    $output .= '
  <table class="table table-striped">
  <thead>
  <tr>
  <th>No.</th>
  <th>Messenger Name</th>
  <th>Address</th>
  <th>Prepared</th>
  <th>Date</th>
  <th>Action</th>
  </tr>
  </thead>
  <tbody>
  ';
    while($row = mysqli_fetch_array($result)) {  
          $count++; 
    $recid = $row["id"];
    $messengername = $row["dispatch_control_messenger_name"];
    $address = $row["dispatch_control_messenger_address"];
    $prepared = $row["dispatch_control_messenger_prepared"];
    $date = $row["dispatch_control_messenger_date"];
    $newDate = date("F d, Y", strtotime($date));
    $newDateUpdate = date("m/d/Y", strtotime($date));

      $output .= '
    <tr>
    <td>'.$count.'</td>        
    <td>'.$messengername.'</td>
    <td>'.$address.'</td>
    <td>'.$prepared.'</td>
    <td>'.$newDate.'</td>
    <td><a href="account.php?id='.$log_id.'&dispatchcontrol=focus&select=view&recordidview='.$recid.'" id="'.$recid.'">View Record</a> | <a target="_blank" href="dispatchcontrol.php?id='.$log_id.'&messengerid='.$recid.'" id="'.$recid.'">Print Receipt</a> | <a target="_blank" href="dispatchproof.php?id='.$log_id.'&messengerid='.$recid.'" id="'.$recid.'">Print Proof</a> | <a href="#" id="'.$recid.'" onclick="openDispatchControlMessengerUpdateDialog(\''.$recid.'\',\''.$messengername.'\',\''.$address.'\',\''.$prepared.'\',\''.$newDateUpdate.'\')">Edit</a> | <a href="#" id="'.$recid.'" onclick="openDispatchControlMessengerDeleteDialog(\''.$recid.'\')">Delete</a></td>
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