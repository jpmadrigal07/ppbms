<?php

if(isset($_POST["updateemail"])) {

  include_once("../includes/db_conn.php");
  include_once("../includes/loginstatus.php");

  $email = $_POST['updateemail'];
  $password = $_POST['password'];

  $sql = "SELECT * FROM ppbms_user WHERE id='$log_id' LIMIT 1";
  $query = mysqli_query($db_conn, $sql);
  $check = mysqli_num_rows($query);

  if($check > 0) { 

    while ($row = mysqli_fetch_array($query, MYSQLI_ASSOC)) {
      $emaildb = $row["user_email"];
      $passworddb = $row["user_pass"];
    }

    if($email == $emaildb && $password == $passworddb) {
      echo "samedata";
    } else {
      $sql = "UPDATE ppbms_user SET user_email='$email', user_pass='$password' WHERE id='$log_id' LIMIT 1";
      $query = mysqli_query($db_conn, $sql);
      echo "successinsert";
    }

  } else {

    echo "error";

  }

  mysqli_close($db_conn);

  mysqli_close($db_conn);

  exit();

}

if(isset($_POST["updatemasterlistrid"])) {

  include_once("../includes/db_conn.php");
  include_once("../includes/loginstatus.php");

  $rid = $_POST['updatemasterlistrid'] == "empty" ? '' : $_POST['updatemasterlistrid'];
  $sender = $_POST['updatemasterlistsender'] == "empty" ? '' : $_POST['updatemasterlistsender'];
  $deltype = $_POST['updatemasterlistdeltype'] == "empty" ? '' : $_POST['updatemasterlistdeltype'];
  $pud = $_POST['updatemasterlistpud'] == "empty" ? '' : $_POST['updatemasterlistpud'];
  $month = $_POST['updatemasterlistmonth'] == "empty" ? '' : $_POST['updatemasterlistmonth'];
  $jobnum = $_POST['updatemasterlistjobnum'] == "empty" ? '' : $_POST['updatemasterlistjobnum'];
  $checklist = $_POST['updatemasterlistchecklist'] == "empty" ? '' : $_POST['updatemasterlistchecklist'];
  $filename = $_POST['updatemasterlistfilename'] == "empty" ? '' : $_POST['updatemasterlistfilename'];
  $seqnum = $_POST['updatemasterlistseqnum'] == "empty" ? '' : $_POST['updatemasterlistseqnum'];
  $cyclecode = $_POST['updatemasterlistcyclecode'] == "empty" ? '' : $_POST['updatemasterlistcyclecode'];
  $qty = $_POST['updatemasterlistqty'] == "empty" ? '' : $_POST['updatemasterlistqty'];
  $address = $_POST['updatemasterlistaddress'] == "empty" ? '' : $_POST['updatemasterlistaddress'];
  $area = $_POST['updatemasterlistarea'] == "empty" ? '' : $_POST['updatemasterlistarea'];
  $subs = $_POST['updatemasterlistsubs'] == "empty" ? '' : $_POST['updatemasterlistsubs'];
  $barcode = $_POST['updatemasterlistbarcode'] == "empty" ? '' : $_POST['updatemasterlistbarcode'];
  $acctnum = $_POST['updatemasterlistacctnum'] == "empty" ? '' : $_POST['updatemasterlistacctnum'];
  $receive = $_POST['updatemasterlistreceive'] == "empty" ? '0000-00-00 00:00:00' : $_POST['updatemasterlistreceive'];
  $receiveby = $_POST['updatemasterlistreceiveby'] == "empty" ? '' : $_POST['updatemasterlistreceiveby'];
  $relation = $_POST['updatemasterlistrelation'] == "empty" ? '' : $_POST['updatemasterlistrelation'];
  $messenger = $_POST['updatemasterlistmessenger'] == "empty" ? '' : $_POST['updatemasterlistmessenger'];
  $status = $_POST['updatemasterliststatus'] == "empty" ? '' : $_POST['updatemasterliststatus'];
  $reasonrts = $_POST['updatemasterlistreasonrts'] == "empty" ? '' : $_POST['updatemasterlistreasonrts'];
  $remarks = $_POST['updatemasterlistreasonremarks'] == "empty" ? '' : $_POST['updatemasterlistreasonremarks'];
  $reported = $_POST['updatemasterlistreasonreported'] == "empty" ? '0000-00-00 00:00:00' : $_POST['updatemasterlistreasonreported'];

  $address = mysqli_real_escape_string($db_conn, $address);

  if($receive != "0000-00-00 00:00:00") {
    $receive = strtotime($receive);
    $receive = date('Y-m-d H:i:s', $receive);
  }

  if($reported != "0000-00-00 00:00:00") {
    $reported = strtotime($reported);
    $reported = date('Y-m-d H:i:s', $reported);
  }

  $sql = "SELECT * FROM ppbms_record WHERE id='$rid' AND record_status_status = '1' LIMIT 1";
  $query = mysqli_query($db_conn, $sql);
  $check = mysqli_num_rows($query);

  if($check > 0) { 

    $sql = "UPDATE ppbms_record SET record_sender='$sender', record_deltype='$deltype', record_pud='$pud', record_month='$month', record_job_num='$jobnum', record_check_list='$checklist', record_file_name='$filename', record_seq_num='$seqnum', record_cycle_code='$cyclecode', record_qty='$qty', record_address='$address', record_area='$area', record_subs_name='$subs', record_bar_code='$barcode', record_acct_num='$acctnum', record_date_receive='$receive', record_receive_by='$receiveby', record_relation='$relation', record_messenger='$messenger', record_status='$status', record_reason_rts='$reasonrts', record_remarks='$remarks', record_date_reported='$reported' WHERE id='$rid' LIMIT 1";
    $query = mysqli_query($db_conn, $sql);
    echo "successupdate";

  } else {

    echo "error";

  }

  mysqli_close($db_conn);

  exit();

}


if(isset($_POST["addbarcodemiddletextcode"])) {

  include_once("../includes/db_conn.php");
  include_once("../includes/loginstatus.php");

  $code = $_POST['addbarcodemiddletextcode'];

  $sql = "SELECT * FROM ppbms_barcode_middle_text WHERE barcode_middle_text_code='$code' LIMIT 1";
  $query = mysqli_query($db_conn, $sql);
  $check = mysqli_num_rows($query);

  if($check > 0) { 

    echo "samedata";

  } else {

    $sql_code = "INSERT INTO ppbms_barcode_middle_text (barcode_middle_text_code, barcode_middle_text_status)
    VALUES ('$code','1')";
    $query = mysqli_query($db_conn, $sql_code);
    echo "successinsert";

  }

  mysqli_close($db_conn);

  exit();

}

if(isset($_POST["updatebarcodemiddletextid"])) {

  include_once("../includes/db_conn.php");
  include_once("../includes/loginstatus.php");

  $rid = $_POST['updatebarcodemiddletextid'];
  $code = $_POST['updatebarcodemiddletextcode'];

  $sql = "UPDATE ppbms_barcode_middle_text SET barcode_middle_text_code='$code' WHERE id='$rid' LIMIT 1";
  $query = mysqli_query($db_conn, $sql);
  echo "successupdate";

  mysqli_close($db_conn);

  exit();

}

if(isset($_POST["updateareapriceid"])) {

  include_once("../includes/db_conn.php");
  include_once("../includes/loginstatus.php");

  $rid = $_POST['updateareapriceid'];
  $price = $_POST['updateareapriceprice'];

  $sql = "UPDATE ppbms_area_price SET area_price_price='$price' WHERE id='$rid' LIMIT 1";
  $query = mysqli_query($db_conn, $sql);
  echo "successupdate";

  mysqli_close($db_conn);

  exit();

}

if(isset($_POST["deletebarcodemiddletextid"])) {

  include_once("../includes/db_conn.php");
  include_once("../includes/loginstatus.php");

  $rid = $_POST['deletebarcodemiddletextid'];

  $sql = "UPDATE ppbms_barcode_middle_text SET barcode_middle_text_status='2' WHERE id='$rid' LIMIT 1";
  $query = mysqli_query($db_conn, $sql);
  echo "successupdate";

  mysqli_close($db_conn);

  exit();

}

if(isset($_POST["deletemasterlistitemid"])) {

  include_once("../includes/db_conn.php");
  include_once("../includes/loginstatus.php");

  $rid = $_POST['deletemasterlistitemid'];

  $sql = "UPDATE ppbms_record SET record_status_status='2' WHERE id='$rid' LIMIT 1";
  $query = mysqli_query($db_conn, $sql);
  echo "successupdate";

  mysqli_close($db_conn);

  exit();

}

if(isset($_POST["adddispatchcontrolmessengername"])) {

  include_once("../includes/db_conn.php");
  include_once("../includes/loginstatus.php");

  $name = $_POST['adddispatchcontrolmessengername'];
  $address = $_POST['adddispatchcontrolmessengeraddress'];
  $prepared = $_POST['adddispatchcontrolmessengerprepared'];
  $date = $_POST['adddispatchcontrolmessengerdate'];
  $newDate = date("Y-m-d H:i:s", strtotime($date));

  $sql = "SELECT * FROM ppbms_dispatch_control_messenger WHERE dispatch_control_messenger_name='$name' AND dispatch_control_messenger_address='$address' AND dispatch_control_messenger_prepared='$address' AND dispatch_control_messenger_date='$address' AND dispatch_control_messenger_status = '1' LIMIT 1";
  $query = mysqli_query($db_conn, $sql);
  $check = mysqli_num_rows($query);

  if($check > 0) { 

    echo "samedata";

  } else {

    $sql_code = "INSERT INTO ppbms_dispatch_control_messenger (dispatch_control_messenger_name, dispatch_control_messenger_address, dispatch_control_messenger_prepared, dispatch_control_messenger_date, dispatch_control_messenger_status)
    VALUES ('$name','$address','$prepared','$newDate','1')";
    $query = mysqli_query($db_conn, $sql_code);
    echo "successinsert";

  }

  mysqli_close($db_conn);

  exit();

}

if(isset($_POST["updatedispatchcontrolmessengerid"])) {

  include_once("../includes/db_conn.php");
  include_once("../includes/loginstatus.php");

  $rid = $_POST['updatedispatchcontrolmessengerid'];
  $name = $_POST['updatedispatchcontrolmessengername'];
  $address = $_POST['updatedispatchcontrolmessengeaddress'];
  $prepared = $_POST['updatedispatchcontrolmessengerprepared'];
  $date = $_POST['updatedispatchcontrolmessengerdate'];
  $newDate = date("Y-m-d H:i:s", strtotime($date));

  $sql = "UPDATE ppbms_dispatch_control_messenger SET dispatch_control_messenger_name='$name', dispatch_control_messenger_address='$address', dispatch_control_messenger_prepared='$prepared', dispatch_control_messenger_date='$newDate' WHERE id='$rid' LIMIT 1";
  $query = mysqli_query($db_conn, $sql);
  echo "successupdate";

  mysqli_close($db_conn);

  exit();

}

if(isset($_POST["deletedispatchcontrolmessengerid"])) {

  include_once("../includes/db_conn.php");
  include_once("../includes/loginstatus.php");

  $rid = $_POST['deletedispatchcontrolmessengerid'];

  $sql = "UPDATE ppbms_dispatch_control_messenger SET dispatch_control_messenger_status='2' WHERE id='$rid' LIMIT 1";
  $query = mysqli_query($db_conn, $sql);
  echo "successupdate";

  mysqli_close($db_conn);

  exit();

}

if(isset($_POST["adddispatchcontroldatamessengerid"])) {

  include_once("../includes/db_conn.php");
  include_once("../includes/loginstatus.php");

  $rid = $_POST['adddispatchcontroldatamessengerid'];
  $cyclecode = $_POST['adddispatchcontroldatacyclecode'];
  $date = $_POST['adddispatchcontroldatapickupdate'];
  $newDate1 = date("Y-m-d H:i:s", strtotime($date));
  $newDate2 = date("m/d/y", strtotime($date));
  $newDate3 = date("n/j/y", strtotime($date));
  $newDate4 = date("m/d/Y", strtotime($date));
  $newDate5 = date("n/j/Y", strtotime($date));
  $sender = $_POST['adddispatchcontroldatasender'];
  $deltype = $_POST['adddispatchcontroldatadeltype'];

  $sql = "SELECT id FROM ppbms_dispatch_control_data WHERE dispatch_control_messenger_id='$rid' AND dispatch_control_data_cycle_code='$cyclecode' AND dispatch_control_data_pickup_date='$newDate1' AND dispatch_control_data_sender='$sender' AND dispatch_control_data_del_type = '$deltype' AND dispatch_control_data_status = '1' LIMIT 1";
  $query = mysqli_query($db_conn, $sql);
  $check = mysqli_num_rows($query);

  $sql_cycle_code = "SELECT id FROM ppbms_record WHERE record_cycle_code='$cyclecode' AND (record_pud='$newDate2' OR record_pud='$newDate3' OR record_pud='$newDate4' OR record_pud='$newDate5') AND record_sender='$sender' AND record_deltype='$deltype' LIMIT 1";
  $query_cycle_code = mysqli_query($db_conn, $sql_cycle_code);
  $check_cycle_code = mysqli_num_rows($query_cycle_code);

  if($check > 0) { 

    echo "samedata";

  } else {

    if($check_cycle_code > 0) {

      $sql_code = "INSERT INTO ppbms_dispatch_control_data (dispatch_control_messenger_id, dispatch_control_data_cycle_code, dispatch_control_data_pickup_date, dispatch_control_data_sender, dispatch_control_data_del_type, dispatch_control_data_status)
      VALUES ('$rid','$cyclecode','$newDate1','$sender','$deltype','1')";
      $query = mysqli_query($db_conn, $sql_code);
      echo "successinsert";

    } else {

      echo "datanotexists";

    }

  }

  mysqli_close($db_conn);

  exit();

}

if(isset($_POST["deleteencodelistsid"])) {

  include_once("../includes/db_conn.php");
  include_once("../includes/loginstatus.php");

  $rid = $_POST['deleteencodelistsid'];

  $sql = "UPDATE ppbms_encode_list SET encode_lists_status='2' WHERE id='$rid' LIMIT 1";
  $query = mysqli_query($db_conn, $sql);

  if($query) {
    $sql1 = "UPDATE ppbms_record SET record_status_status='2' WHERE encode_lists_id='$rid'";
    $query1 = mysqli_query($db_conn, $sql1);
  }
  echo "successupdate";

  mysqli_close($db_conn);

  exit();

}

if(isset($_POST["updaterecordbarcode"])) {

  include_once("../includes/db_conn.php");
  include_once("../includes/loginstatus.php");

  $barcode = $_POST['updaterecordbarcode'];
  $daterecieve = $_POST['updaterecorddaterecieve'];
  $recieveby = $_POST['updaterecordrecieveby'];
  $relation = $_POST['updaterecordrelation'];
  $messenger = $_POST['updaterecordmessenger'];
  $status = $_POST['updaterecordstatus'];
  $reasonrts = $_POST['updaterecordreasonrts'];
  $remarks = $_POST['updaterecordremarks'];
  $datereported = $_POST['updaterecorddatereported'];
  $newDate1 = date("Y-m-d H:i:s", strtotime($daterecieve));
  $newDate2 = date("Y-m-d H:i:s", strtotime($datereported));

  $sql = "UPDATE ppbms_record SET record_date_receive='$newDate1', record_receive_by='$recieveby',   record_relation='$relation', record_messenger='$messenger', record_status='$status', record_reason_rts='$reasonrts', record_remarks='$remarks', record_date_reported='$newDate2' WHERE record_bar_code='$barcode' LIMIT 1";
  $query = mysqli_query($db_conn, $sql);
  echo "successupdate";

  mysqli_close($db_conn);

  exit();

}

if(isset($_POST["qname"])) {

  include_once("../includes/db_conn.php");
  include_once("../includes/loginstatus.php");

  $name = $_POST['qname'];

  $sql = "SELECT * FROM oes_question_list WHERE q_list_name='$name' LIMIT 1";
  $query = mysqli_query($db_conn, $sql);
  $check = mysqli_num_rows($query);

  if($check > 0) { 

    echo "samedata";

  } else {

    $sql_user = "INSERT INTO oes_question_list (q_list_name, q_list_date_created, q_list_status)
    VALUES ('$name', NOW(), '1')";
    $query = mysqli_query($db_conn, $sql_user);
    echo "successinsert";

  }

  mysqli_close($db_conn);

  exit();

}

if(isset($_POST["aname"])) {

  include_once("../includes/db_conn.php");
  include_once("../includes/loginstatus.php");

  $name = $_POST['aname'];

  $sql = "SELECT * FROM oes_answer_list WHERE a_list_name='$name' LIMIT 1";
  $query = mysqli_query($db_conn, $sql);
  $check = mysqli_num_rows($query);

  if($check > 0) { 

    echo "samedata";

  } else {

    $sql_user = "INSERT INTO oes_answer_list (a_list_name, a_list_date_created, a_list_status)
    VALUES ('$name', NOW(), '1')";
    $query = mysqli_query($db_conn, $sql_user);
    echo "successinsert";

  }

  mysqli_close($db_conn);

  exit();

}

?>