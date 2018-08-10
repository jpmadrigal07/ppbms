<?php  
include_once("../includes/loginstatus.php");

if(isset($_POST["cyclecode"]) && isset($_POST["messengerid"]) && isset($_POST["search"])) {

  $search = $_POST["search"];
  $cyclecode = $_POST["cyclecode"];
  $messengerid = $_POST["messengerid"];

  $sql = "SELECT * FROM ppbms_record WHERE record_status_status = '1' AND record_bar_code = '".$search."' AND record_cycle_code = '".$cyclecode."' LIMIT 1";  
  $result = mysqli_query($db_conn, $sql);  

  if(mysqli_num_rows($result) > 0) {  
    $sql = "UPDATE ppbms_record SET messenger_id='$messengerid' WHERE record_bar_code = '".$search."' AND record_cycle_code = '".$cyclecode."' LIMIT 1";
    $query = mysqli_query($db_conn, $sql);
      echo "successupdate";
  } else {
    echo "failedupdate";
  }

  mysqli_close($db_conn);

  exit(); 

}

?>  