<?php  
include_once("../includes/loginstatus.php");

$output = '';
$count = 0; 

$date = strtotime($_POST["search"]);

$newDate = date('Y-m-d', $date);

$sql = "SELECT id, encode_file_name, encode_lists_date, encode_lists_date FROM ppbms_encode_list WHERE cast(encode_lists_date AS date) = '$newDate' AND encode_lists_status = '1' ORDER BY id DESC"; 
$result = mysqli_query($db_conn, $sql);  


if($_POST["search"] == "") {
  $output .= '
  <div id="pagination_controls_encode_lists_1"></div>
  <div id="record-search-result-encode-lists"></div>
  <div id="pagination_controls_encode_lists_2"></div>
  <script type="text/javascript">
  request_page_encode_lists(1);
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
    <th>File Name</th>
    <th>Date Imported</th>
    <th>Assigned Count</th>
    <th>Unassigned Count</th>
    <th>Total Count</th>
    <th>Action</th>
    </tr>
    </thead>
    <tbody>
    ';
    while($row = mysqli_fetch_array($result)) {  
      $count++; 
       $recid = $row["id"];
       $filename = $row["encode_file_name"];
       $date = $row["encode_lists_date"];
       $newDate = date("M d, Y - g:i A", strtotime($date));
       $newDateUpdate = date("m/d/Y H:i A", strtotime($date));

       $sql_unassined = "SELECT id FROM ppbms_record WHERE encode_lists_id = '$recid' AND messenger_id = '0' AND record_status_status = '1'";
       $query_unassined = mysqli_query($db_conn, $sql_unassined);
       $unassined_count = mysqli_num_rows($query_unassined);

       $sql_assined = "SELECT id FROM ppbms_record WHERE encode_lists_id = '$recid' AND messenger_id != '0' AND record_status_status = '1'";
       $query_assined = mysqli_query($db_conn, $sql_assined);
       $assined_count = mysqli_num_rows($query_assined);

       $sql_total = "SELECT id FROM ppbms_record WHERE encode_lists_id = '$recid' AND record_status_status = '1'";
       $query_total= mysqli_query($db_conn, $sql_total);
       $total_count = mysqli_num_rows($query_total);

      $output .= '
      <tr>
     <td>'.$count.'</td>
     <td>'.$filename.'</td>         
     <td>'.$newDate.'</td>
     <td>'.$assined_count.'</td>
     <td><a id="'.$recid.'" href="account.php?id='.$log_id.'&masterlists=focus&select=view&encodeid='.$recid.'&check=unassigned">'.$unassined_count.'</a></td>
     <td>'.$total_count.'</td>
     <td><a id="'.$recid.'" href="account.php?id='.$log_id.'&masterlists=focus&select=view&encodeid='.$recid.'">View Record</a> | <a target="_blank" id="'.$recid.'" href="receipt.php?id='.$log_id.'&encodeid='.$recid.'">Print Receipt</a> | <a href="#" id="'.$recid.'" onclick="openEncodeListsDeleteDialog(\''.$recid.'\')">Delete</a></td>
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