<?php  
include_once("../includes/loginstatus.php");

if(isset($_POST["check"]) && $_POST["check"] == "1") {

  $search = $_POST["search"];

  $sql = "SELECT id FROM ppbms_record WHERE record_status_status = '1' AND record_bar_code = '".$_POST["search"]."' LIMIT 1";  
  $result = mysqli_query($db_conn, $sql);  

  if(mysqli_num_rows($result) > 0) {  
    echo "true";
  } else {
    echo "false";
  }

  mysqli_close($db_conn);

  exit(); 

} else if(isset($_POST["check"]) && $_POST["check"] == "2") {

  $sql = "SELECT record_date_receive FROM ppbms_record WHERE record_status_status = '1' AND record_bar_code = '".$_POST["search"]."' LIMIT 1";  
  $result = mysqli_query($db_conn, $sql);  

  if(mysqli_num_rows($result) > 0) {  
    while($row = mysqli_fetch_array($result)) {  
      $date = $row["record_date_receive"];
      if($date == "0000-00-00 00:00:00") {
        echo "";
      } else {
        echo date("m/d/Y", strtotime($date));
      }
    }
  } else {
    echo "";
  }

  mysqli_close($db_conn);

  exit(); 

} else if(isset($_POST["check"]) && $_POST["check"] == "3") {

  $sql = "SELECT record_receive_by FROM ppbms_record WHERE record_status_status = '1' AND record_bar_code = '".$_POST["search"]."' LIMIT 1";  
  $result = mysqli_query($db_conn, $sql);  

  if(mysqli_num_rows($result) > 0) {  
    while($row = mysqli_fetch_array($result)) {  
      echo $row["record_receive_by"];
    }
  } else {
    echo "";
  }

  mysqli_close($db_conn);

  exit(); 

} else if(isset($_POST["check"]) && $_POST["check"] == "4") {

  $sql = "SELECT record_relation FROM ppbms_record WHERE record_status_status = '1' AND record_bar_code = '".$_POST["search"]."' LIMIT 1";  
  $result = mysqli_query($db_conn, $sql);  

  if(mysqli_num_rows($result) > 0) {  
    while($row = mysqli_fetch_array($result)) {  
      echo $row["record_relation"];
    }
  } else {
    echo "";
  }

  mysqli_close($db_conn);

  exit(); 

} else if(isset($_POST["check"]) && $_POST["check"] == "5") {

  $sql = "SELECT record_messenger FROM ppbms_record WHERE record_status_status = '1' AND record_bar_code = '".$_POST["search"]."' LIMIT 1";  
  $result = mysqli_query($db_conn, $sql);  

  if(mysqli_num_rows($result) > 0) {  
    while($row = mysqli_fetch_array($result)) {  
      echo $row["record_messenger"];
    }
  } else {
    echo "";
  }

  mysqli_close($db_conn);

  exit(); 

} else if(isset($_POST["check"]) && $_POST["check"] == "6") {

  $sql = "SELECT record_status FROM ppbms_record WHERE record_status_status = '1' AND record_bar_code = '".$_POST["search"]."' LIMIT 1";  
  $result = mysqli_query($db_conn, $sql);  

  if(mysqli_num_rows($result) > 0) {  
    while($row = mysqli_fetch_array($result)) {  
      echo $row["record_status"];
    }
  } else {
    echo "";
  }

  mysqli_close($db_conn);

  exit(); 

} else if(isset($_POST["check"]) && $_POST["check"] == "7") {

  $sql = "SELECT record_reason_rts FROM ppbms_record WHERE record_status_status = '1' AND record_bar_code = '".$_POST["search"]."' LIMIT 1";  
  $result = mysqli_query($db_conn, $sql);  

  if(mysqli_num_rows($result) > 0) {  
    while($row = mysqli_fetch_array($result)) {  
      echo $row["record_reason_rts"];
    }
  } else {
    echo "";
  }

  mysqli_close($db_conn);

  exit(); 

} else if(isset($_POST["check"]) && $_POST["check"] == "8") {

  $sql = "SELECT record_remarks FROM ppbms_record WHERE record_status_status = '1' AND record_bar_code = '".$_POST["search"]."' LIMIT 1";  
  $result = mysqli_query($db_conn, $sql);  

  if(mysqli_num_rows($result) > 0) {  
    while($row = mysqli_fetch_array($result)) {  
      echo $row["record_remarks"];
    }
  } else {
    echo "";
  }

  mysqli_close($db_conn);

  exit(); 

} else if(isset($_POST["check"]) && $_POST["check"] == "9") {

  $sql = "SELECT record_date_reported FROM ppbms_record WHERE record_status_status = '1' AND record_bar_code = '".$_POST["search"]."' LIMIT 1";  
  $result = mysqli_query($db_conn, $sql);  

  if(mysqli_num_rows($result) > 0) {  
    while($row = mysqli_fetch_array($result)) {  
      $date = $row["record_date_reported"];
      if($date == "0000-00-00 00:00:00") {
        echo "";
      } else {
        echo date("m/d/Y", strtotime($date));
      }
    }
  } else {
    echo "";
  }

  mysqli_close($db_conn);

  exit(); 

}

?>  