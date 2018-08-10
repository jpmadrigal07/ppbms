<?php
include_once("includes/loginstatus.php");
if (!isset($_SESSION["userid"])) {
  header("location: index.php");
  exit();
}
// VARIABLES
$id = "";
$accname = "";
$email = "";
$password = "";
$avatar = "";
$ip = "";
$datereg = "";
$lastlogin = "";
$level = "";
$statusverify = "";
$townname = "";
$pheading = "";
$pbody = "";
$classactivedashboard = "";
$classactivesettings = "";
$classactivemasterlists = "";
$classactiveencodemasterlists = "";
$echo = "";
$messengerid = "";
$encodeid = "0";
$messengername = "";
$address = "";
$prepared = "";
$dispatchcontrolnewdate = "";
$newPUD1 = "";
$deltype1 = "";
$recid1 = "";
$cyclecode1 = "";
$pickupdate1 = "";
$sender1 = "";
$qty1 = 0;
$newPUD2 = "";
$deltype2 = "";
$recid2 = "";
$cyclecode2 = "";
$pickupdate2 = "";
$sender2 = "";
$qty2 = 0;
$newPUD3 = "";
$recid3 = "";
$cyclecode3 = "";
$pickupdate3 = "";
$sender3 = "";
$deltype3 = "";
$qty3 = 0;
$newPUD4 = "";
$recid4 = "";
$cyclecode4 = "";
$pickupdate4 = "";
$sender4 = "";
$deltype4 = "";
$qty4 = 0;
$newPUD5 = "";
$recid5 = "";
$cyclecode5 = "";
$pickupdate5 = "";
$sender5 = "";
$deltype5 = "";
$qty5 = 0;
$newPUD6 = "";
$recid6 = "";
$cyclecode6 = "";
$pickupdate6 = "";
$sender6 = "";
$deltype6 = "";
$qty6 = 0;
$newPUD7 = "";
$recid7 = "";
$cyclecode7 = "";
$pickupdate7 = "";
$sender7 = "";
$deltype7 = "";
$qty7 = 0;
$newPUD8 = "";
$recid8 = "";
$cyclecode8 = "";
$pickupdate8 = "";
$sender8 = "";
$deltype8 = "";
$qty8 = 0;
$newPUD9 = "";
$recid9 = "";
$cyclecode9 = "";
$pickupdate9 = "";
$sender9 = "";
$deltype9 = "";
$qty9 = 0;
$newPUD10 = "";
$recid10 = "";
$cyclecode10 = "";
$pickupdate10 = "";
$sender10 = "";
$deltype10 = "";
$qty10 = 0;
$newPUD11 = "";
$recid11 = "";
$cyclecode11 = "";
$pickupdate11 = "";
$sender11 = "";
$deltype11 = "";
$qty11 = 0;
$newPUD12 = "";
$recid12 = "";
$cyclecode12 = "";
$pickupdate12 = "";
$sender12 = "";
$deltype12 = "";
$qty12 = 0;
$newPUD13 = "";
$recid13 = "";
$cyclecode13 = "";
$pickupdate13 = "";
$sender13 = "";
$deltype13 = "";
$qty13 = 0;
$newPUD14 = "";
$cyclecode14 = "";
$pickupdate14 = "";
$sender14 = "";
$deltype14 = "";
$qty14 = 0;
$newPUD15 = "";
$cyclecode15 = "";
$pickupdate15 = "";
$sender15 = "";
$deltype15 = "";
$qty15 = 0;
$newPUD16 = "";
$cyclecode16 = "";
$pickupdate16 = "";
$sender16 = "";
$deltype16 = "";
$qty16 = 0;
$newPUD17 = "";
$cyclecode17 = "";
$pickupdate17 = "";
$sender17 = "";
$deltype17 = "";
$qty17 = 0;
$newPUD18 = "";
$cyclecode18 = "";
$pickupdate18 = "";
$sender18 = "";
$deltype18 = "";
$qty18 = 0;
$newPUD19 = "";
$cyclecode19 = "";
$pickupdate19 = "";
$sender19 = "";
$deltype19 = "";
$qty19 = 0;
$newPUD20 = "";
$cyclecode20 = "";
$pickupdate20 = "";
$sender20 = "";
$deltype20 = "";
$qty20 = 0;
$newPUD21 = "";
$cyclecode21 = "";
$pickupdate21 = "";
$sender21 = "";
$deltype21 = "";
$qty21 = 0;
$newPUD22 = "";
$cyclecode22 = "";
$pickupdate22 = "";
$sender22 = "";
$deltype22 = "";
$qty22 = 0;
$newPUD23 = "";
$cyclecode23 = "";
$pickupdate23 = "";
$sender23 = "";
$deltype23 = "";
$qty23 = 0;
$newPUD24 = "";
$cyclecode24 = "";
$pickupdate24 = "";
$sender24 = "";
$deltype24 = "";
$qty24 = 0;
$grandTotal = 0;

// GET USER ID
if(isset($_GET["id"])){
  $id = preg_replace('#[^a-z0-9-]#i', '', $_GET['id']);
} else {
  header("location: index.php");
  exit(); 
}

// SELECT USER INFO
$sql_user = "SELECT * FROM ppbms_user WHERE id='$id' LIMIT 1";
$user_query = mysqli_query($db_conn, $sql_user);
while ($row = mysqli_fetch_array($user_query, MYSQLI_ASSOC)) {
  $id = $row["id"];
  $email = $row["user_email"];
  $password = $row["user_pass"];
  $ip = $row["user_ip"];
  $datereg = $row["user_date_created"];
  $lastlogin = $row["user_last_login"];
  $level = $row["user_level"];
}

if($level == "1") {
    $accname = "Administrator";
}

// CHECK TO SEE IF THE VIEWER IS THE ACCOUNT USER
$isOwner = "No";

if($id == $log_id && $user_ok == true){
  $isOwner = "Yes";
}

if(isset($_GET["messengerid"])){
  $messengerid = $_GET['messengerid'];
  $sql_dispatch_control = "SELECT dispatch_control_messenger_name, dispatch_control_messenger_address, dispatch_control_messenger_prepared, dispatch_control_messenger_date FROM ppbms_dispatch_control_messenger WHERE dispatch_control_messenger_status = '1' AND id = '$messengerid' LIMIT 1";
    $query_dispatch_control = mysqli_query($db_conn, $sql_dispatch_control);
    while($row_dispatch_control = mysqli_fetch_array($query_dispatch_control)) {  
        $messengername = $row_dispatch_control["dispatch_control_messenger_name"];
        $address = $row_dispatch_control["dispatch_control_messenger_address"];
        $prepared = $row_dispatch_control["dispatch_control_messenger_prepared"];
        $dispatchcontroldate = $row_dispatch_control["dispatch_control_messenger_date"];
        $dispatchcontrolnewdate = date("M d, Y", strtotime($dispatchcontroldate));
    }
  $count = 0;
  $sql = "SELECT id, dispatch_control_data_cycle_code, dispatch_control_data_pickup_date, dispatch_control_data_sender, dispatch_control_data_del_type FROM ppbms_dispatch_control_data WHERE dispatch_control_messenger_id='$messengerid' AND dispatch_control_data_status='1' LIMIT 24";
    $query = mysqli_query($db_conn, $sql);
    while ($row = mysqli_fetch_array($query, MYSQLI_ASSOC)) {
      $count++;
      if($count == 1) {
          $recid1 = $row["id"];
          $cyclecode1 = $row["dispatch_control_data_cycle_code"];
          $pickupdate1 = $row["dispatch_control_data_pickup_date"];
          $sender1 = $row["dispatch_control_data_sender"];
          $deltype1 = $row["dispatch_control_data_del_type"];
          $newPUD1 = date("m/d/y", strtotime($pickupdate1));
          $newPUD12 = date("n/j/y", strtotime($pickupdate1));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender1' AND record_deltype = '$deltype1' AND (record_pud = '$newPUD1' OR record_pud = '$newPUD12') AND record_cycle_code = '$cyclecode1' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty1 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 2) {
          $recid2 = $row["id"];
          $cyclecode2 = $row["dispatch_control_data_cycle_code"];
          $pickupdate2 = $row["dispatch_control_data_pickup_date"];
          $sender2 = $row["dispatch_control_data_sender"];
          $deltype2 = $row["dispatch_control_data_del_type"];
          $newPUD2 = date("m/d/y", strtotime($pickupdate2));
          $newPUD22 = date("n/j/y", strtotime($pickupdate2));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender2' AND record_deltype = '$deltype2' AND (record_pud = '$newPUD2' OR record_pud = '$newPUD22') AND record_cycle_code = '$cyclecode2' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty2 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 3) {
          $recid3 = $row["id"];
          $cyclecode3 = $row["dispatch_control_data_cycle_code"];
          $pickupdate3 = $row["dispatch_control_data_pickup_date"];
          $sender3 = $row["dispatch_control_data_sender"];
          $deltype3 = $row["dispatch_control_data_del_type"];
          $newPUD3 = date("m/d/y", strtotime($pickupdate3));
          $newPUD32 = date("n/j/y", strtotime($pickupdate3));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender3' AND record_deltype = '$deltype3' AND (record_pud = '$newPUD3' OR record_pud = '$newPUD32') AND record_cycle_code = '$cyclecode3' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
         if($query_check_record_receive) {
            $qty3 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 4) {
          $recid4 = $row["id"];
          $cyclecode4 = $row["dispatch_control_data_cycle_code"];
          $pickupdate4 = $row["dispatch_control_data_pickup_date"];
          $sender4 = $row["dispatch_control_data_sender"];
          $deltype4 = $row["dispatch_control_data_del_type"];
          $newPUD4 = date("m/d/y", strtotime($pickupdate4));
          $newPUD42 = date("n/j/y", strtotime($pickupdate4));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender4' AND record_deltype = '$deltype4' AND (record_pud = '$newPUD4' OR record_pud = '$newPUD42') AND record_cycle_code = '$cyclecode4' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty4 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 5) {
          $recid5 = $row["id"];
          $cyclecode5 = $row["dispatch_control_data_cycle_code"];
          $pickupdate5 = $row["dispatch_control_data_pickup_date"];
          $sender5 = $row["dispatch_control_data_sender"];
          $deltype5 = $row["dispatch_control_data_del_type"];
          $newPUD5 = date("m/d/y", strtotime($pickupdate5));
          $newPUD52 = date("n/j/y", strtotime($pickupdate5));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender5' AND record_deltype = '$deltype5' AND (record_pud = '$newPUD5' OR record_pud = '$newPUD52') AND record_cycle_code = '$cyclecode5' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty5 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 6) {
          $recid6 = $row["id"];
          $cyclecode6 = $row["dispatch_control_data_cycle_code"];
          $pickupdate6 = $row["dispatch_control_data_pickup_date"];
          $sender6 = $row["dispatch_control_data_sender"];
          $deltype6 = $row["dispatch_control_data_del_type"];
          $newPUD6 = date("m/d/y", strtotime($pickupdate6));
          $newPUD62 = date("n/j/y", strtotime($pickupdate6));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender6' AND record_deltype = '$deltype6' AND (record_pud = '$newPUD6' OR record_pud = '$newPUD62') AND record_cycle_code = '$cyclecode6' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty6 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 7) {
          $recid7 = $row["id"];
          $cyclecode7 = $row["dispatch_control_data_cycle_code"];
          $pickupdate7 = $row["dispatch_control_data_pickup_date"];
          $sender7 = $row["dispatch_control_data_sender"];
          $deltype7 = $row["dispatch_control_data_del_type"];
          $newPUD7 = date("m/d/y", strtotime($pickupdate7));
          $newPUD72 = date("n/j/y", strtotime($pickupdate7));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender7' AND record_deltype = '$deltype7' AND (record_pud = '$newPUD7' OR record_pud = '$newPUD72') AND record_cycle_code = '$cyclecode7' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty7 = mysqli_num_rows($query_check_record_receive);
          }
      } else if($count == 8) {
          $recid8 = $row["id"];
          $cyclecode8 = $row["dispatch_control_data_cycle_code"];
          $pickupdate8 = $row["dispatch_control_data_pickup_date"];
          $sender8 = $row["dispatch_control_data_sender"];
          $deltype8 = $row["dispatch_control_data_del_type"];
          $newPUD8 = date("m/d/y", strtotime($pickupdate8));
          $newPUD82 = date("n/j/y", strtotime($pickupdate8));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender8' AND record_deltype = '$deltype8' AND (record_pud = '$newPUD8' OR record_pud = '$newPUD82') AND record_cycle_code = '$cyclecode8' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty8 = mysqli_num_rows($query_check_record_receive);
          }
      } else if($count == 9) {
          $recid9 = $row["id"];
          $cyclecode9 = $row["dispatch_control_data_cycle_code"];
          $pickupdate9 = $row["dispatch_control_data_pickup_date"];
          $sender9 = $row["dispatch_control_data_sender"];
          $deltype9 = $row["dispatch_control_data_del_type"];
          $newPUD9 = date("m/d/y", strtotime($pickupdate9));
          $newPUD92 = date("n/j/y", strtotime($pickupdate9));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender9' AND record_deltype = '$deltype9' AND (record_pud = '$newPUD9' OR record_pud = '$newPUD92') AND record_cycle_code = '$cyclecode9' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty9 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 10) {
          $recid10 = $row["id"];
          $cyclecode10 = $row["dispatch_control_data_cycle_code"];
          $pickupdate10 = $row["dispatch_control_data_pickup_date"];
          $sender10 = $row["dispatch_control_data_sender"];
          $deltype10 = $row["dispatch_control_data_del_type"];
          $newPUD10 = date("m/d/y", strtotime($pickupdate10));
          $newPUD102 = date("n/j/y", strtotime($pickupdate10));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender10' AND record_deltype = '$deltype10' AND (record_pud = '$newPUD10' OR record_pud = '$newPUD102') AND record_cycle_code = '$cyclecode10' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty10 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 11) {
          $recid11 = $row["id"];
          $cyclecode11 = $row["dispatch_control_data_cycle_code"];
          $pickupdate11 = $row["dispatch_control_data_pickup_date"];
          $sender11 = $row["dispatch_control_data_sender"];
          $deltype11 = $row["dispatch_control_data_del_type"];
          $newPUD11 = date("m/d/y", strtotime($pickupdate11));
          $newPUD112 = date("n/j/y", strtotime($pickupdate11));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender11' AND record_deltype = '$deltype11' AND (record_pud = '$newPUD11' OR record_pud = '$newPUD112') AND record_cycle_code = '$cyclecode11' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty11 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 12) {
          $recid1 = $row["id"];
          $cyclecode12 = $row["dispatch_control_data_cycle_code"];
          $pickupdate12 = $row["dispatch_control_data_pickup_date"];
          $sender12 = $row["dispatch_control_data_sender"];
          $deltype12 = $row["dispatch_control_data_del_type"];
          $newPUD12 = date("m/d/y", strtotime($pickupdate12));
          $newPUD122 = date("n/j/y", strtotime($pickupdate12));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender12' AND record_deltype = '$deltype12' AND (record_pud = '$newPUD12' OR record_pud = '$newPUD122') AND record_cycle_code = '$cyclecode12' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty12 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 13) {
          $recid13 = $row["id"];
          $cyclecode13 = $row["dispatch_control_data_cycle_code"];
          $pickupdate13 = $row["dispatch_control_data_pickup_date"];
          $sender13 = $row["dispatch_control_data_sender"];
          $deltype13 = $row["dispatch_control_data_del_type"];
          $newPUD13 = date("m/d/y", strtotime($pickupdate13));
          $newPUD132 = date("n/j/y", strtotime($pickupdate13));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender13' AND record_deltype = '$deltype13' AND (record_pud = '$newPUD13' OR record_pud = '$newPUD132') AND record_cycle_code = '$cyclecode13' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty13 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 14) {
          $recid14 = $row["id"];
          $cyclecode14 = $row["dispatch_control_data_cycle_code"];
          $pickupdate14 = $row["dispatch_control_data_pickup_date"];
          $sender14 = $row["dispatch_control_data_sender"];
          $deltype14 = $row["dispatch_control_data_del_type"];
          $newPUD14 = date("m/d/y", strtotime($pickupdate14));
          $newPUD142 = date("n/j/y", strtotime($pickupdate14));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender14' AND record_deltype = '$deltype14' AND (record_pud = '$newPUD14' OR record_pud = '$newPUD142') AND record_cycle_code = '$cyclecode14' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty14 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 15) {
          $recid15 = $row["id"];
          $cyclecode15 = $row["dispatch_control_data_cycle_code"];
          $pickupdate15 = $row["dispatch_control_data_pickup_date"];
          $sender15 = $row["dispatch_control_data_sender"];
          $deltype15 = $row["dispatch_control_data_del_type"];
          $newPUD15 = date("m/d/y", strtotime($pickupdate15));
          $newPUD152 = date("n/j/y", strtotime($pickupdate15));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender15' AND record_deltype = '$deltype15' AND (record_pud = '$newPUD15' OR record_pud = '$newPUD152') AND record_cycle_code = '$cyclecode15' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty15 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 16) {
          $recid16 = $row["id"];
          $cyclecode16 = $row["dispatch_control_data_cycle_code"];
          $pickupdate16 = $row["dispatch_control_data_pickup_date"];
          $sender16 = $row["dispatch_control_data_sender"];
          $deltype16 = $row["dispatch_control_data_del_type"];
          $newPUD16 = date("m/d/y", strtotime($pickupdate16));
          $newPUD162 = date("n/j/y", strtotime($pickupdate16));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender16' AND record_deltype = '$deltype16' AND (record_pud = '$newPUD16' OR record_pud = '$newPUD162') AND record_cycle_code = '$cyclecode16' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty16 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 17) {
          $recid17 = $row["id"];
          $cyclecode17 = $row["dispatch_control_data_cycle_code"];
          $pickupdate17 = $row["dispatch_control_data_pickup_date"];
          $sender17 = $row["dispatch_control_data_sender"];
          $deltype17 = $row["dispatch_control_data_del_type"];
          $newPUD17 = date("m/d/y", strtotime($pickupdate17));
          $newPUD172 = date("n/j/y", strtotime($pickupdate17));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender17' AND record_deltype = '$deltype17' AND (record_pud = '$newPUD17' OR record_pud = '$newPUD172') AND record_cycle_code = '$cyclecode17' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty17 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 18) {
          $recid18 = $row["id"];
          $cyclecode18 = $row["dispatch_control_data_cycle_code"];
          $pickupdate18 = $row["dispatch_control_data_pickup_date"];
          $sender18 = $row["dispatch_control_data_sender"];
          $deltype18 = $row["dispatch_control_data_del_type"];
          $newPUD18 = date("m/d/y", strtotime($pickupdate18));
          $newPUD182 = date("n/j/y", strtotime($pickupdate18));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender18' AND record_deltype = '$deltype18' AND (record_pud = '$newPUD18' OR record_pud = '$newPUD182') AND record_cycle_code = '$cyclecode18' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty18 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 19) {
          $recid19 = $row["id"];
          $cyclecode19 = $row["dispatch_control_data_cycle_code"];
          $pickupdate19 = $row["dispatch_control_data_pickup_date"];
          $sender19 = $row["dispatch_control_data_sender"];
          $deltype19 = $row["dispatch_control_data_del_type"];
          $newPUD19 = date("m/d/y", strtotime($pickupdate19));
          $newPUD192 = date("n/j/y", strtotime($pickupdate19));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender19' AND record_deltype = '$deltype19' AND (record_pud = '$newPUD19' OR record_pud = '$newPUD192') AND record_cycle_code = '$cyclecode19' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty19 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 20) {
          $recid20 = $row["id"];
          $cyclecode20 = $row["dispatch_control_data_cycle_code"];
          $pickupdate20 = $row["dispatch_control_data_pickup_date"];
          $sender20 = $row["dispatch_control_data_sender"];
          $deltype20 = $row["dispatch_control_data_del_type"];
          $newPUD20 = date("m/d/y", strtotime($pickupdate20));
          $newPUD202 = date("n/j/y", strtotime($pickupdate20));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender20' AND record_deltype = '$deltype20' AND (record_pud = '$newPUD20' OR record_pud = '$newPUD202') AND record_cycle_code = '$cyclecode20' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty20 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 21) {
          $recid21 = $row["id"];
          $cyclecode21 = $row["dispatch_control_data_cycle_code"];
          $pickupdate21 = $row["dispatch_control_data_pickup_date"];
          $sender21 = $row["dispatch_control_data_sender"];
          $deltype21 = $row["dispatch_control_data_del_type"];
          $newPUD21 = date("m/d/y", strtotime($pickupdate21));
          $newPUD212 = date("n/j/y", strtotime($pickupdate21));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender21' AND record_deltype = '$deltype21' AND (record_pud = '$newPUD21' OR record_pud = '$newPUD212') AND record_cycle_code = '$cyclecode21' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty21 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 22) {
          $recid22 = $row["id"];
          $cyclecode22 = $row["dispatch_control_data_cycle_code"];
          $pickupdate22 = $row["dispatch_control_data_pickup_date"];
          $sender22 = $row["dispatch_control_data_sender"];
          $deltype22 = $row["dispatch_control_data_del_type"];
          $newPUD22 = date("m/d/y", strtotime($pickupdate22));
          $newPUD222 = date("n/j/y", strtotime($pickupdate22));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender22' AND record_deltype = '$deltype22' AND (record_pud = '$newPUD22' OR record_pud = '$newPUD222') AND record_cycle_code = '$cyclecode22' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty22 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 23) {
          $recid23 = $row["id"];
          $cyclecode23 = $row["dispatch_control_data_cycle_code"];
          $pickupdate23 = $row["dispatch_control_data_pickup_date"];
          $sender23 = $row["dispatch_control_data_sender"];
          $deltype23 = $row["dispatch_control_data_del_type"];
          $newPUD23 = date("m/d/y", strtotime($pickupdate23));
          $newPUD232 = date("n/j/y", strtotime($pickupdate23));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender23' AND record_deltype = '$deltype23' AND (record_pud = '$newPUD23' OR record_pud = '$newPUD232') AND record_cycle_code = '$cyclecode23' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty23 = mysqli_num_rows($query_check_record_receive);
          }

      } else if($count == 24) {
          $recid24 = $row["id"];
          $cyclecode24 = $row["dispatch_control_data_cycle_code"];
          $pickupdate24 = $row["dispatch_control_data_pickup_date"];
          $sender24 = $row["dispatch_control_data_sender"];
          $deltype24 = $row["dispatch_control_data_del_type"];
          $newPUD24 = date("m/d/y", strtotime($pickupdate24));
          $newPUD242 = date("n/j/y", strtotime($pickupdate24));

          $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE record_sender = '$sender24' AND record_deltype = '$deltype24' AND (record_pud = '$newPUD24' OR record_pud = '$newPUD242') AND record_cycle_code = '$cyclecode24' AND messenger_id = '$messengerid' AND record_status_status = '1'";
          $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
          if($query_check_record_receive) {
            $qty24 = mysqli_num_rows($query_check_record_receive);
          }

      }

      $grandTotal = $qty1+$qty2+$qty3+$qty4+$qty5+$qty6+$qty7+$qty8+$qty9+$qty10+$qty11+$qty12+$qty13+$qty14+$qty15+$qty16+$qty17+$qty18+$qty19+$qty20+$qty21+$qty22+$qty23+$qty24;

    }
}

mysqli_close($db_conn);

?>
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- The above 3 meta tags *must* come first in the head; any other head content must come *after* these tags -->
    <title>PPBMS - Dispatch Control</title>
    <!-- FavIcon -->
    <link rel="icon" type="image/png" href="img/fav.png" />
    <!-- Bootstrap -->
    <link href="css/bootstrap.min.css" rel="stylesheet">

    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.3/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->
    <style type="text/css">
    .panel {
      border-radius: 0px;
      border: 1px solid #000000;
    }
    .panel > .panel-heading {
        border-bottom: 1px solid #000000;
        background-color: black !important;
        -webkit-print-color-adjust: exact; 
    }
    table.table-bordered{
        border: 1px solid #000000 !important;
        font-size: 12px;
    }
    table.table-bordered > thead > tr > th{
        border: 1px solid #000000 !important;
        padding: 1px;
    }
    table.table-bordered > tbody > tr > td{
        border: 1px solid #000000 !important;
        padding: 1px;
    }
    </style>
  </head>
  <body>

     <div class="container-fluid">

         <div class="row">
            <div class="col-xs-7 col-sm-7 col-md-7 col-lg-7">
                <h2 style="margin-bottom: 0px; margin-top: 50px;">P.P.B. Messengerial Services</h2>
            </div>
            <div class="col-xs-5 col-sm-5 col-md-5 col-lg-5">
                <p style="float: right; text-align: right; padding-bottom: 0px;">Phase 7 Block 9 Lot 11 Gamet St.,<br>Brgy. San Vicente, Pacita Complex 1,<br>San Pedro, Laguna 4023<br>Tel Nos. : 869-2622 / 478-2822</p>
            </div>
         </div>
         <div class="row">
            <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
                <hr style="border-top: 3px solid #000000; margin-top: 0px; margin-bottom: 0px;">
            </div>
         </div>
         <br>
         <div class="row">
            <div class="col-xs-6 col-sm-6 col-md-6 col-lg-6">
                 Messenger Name: <b><?php echo $messengername; ?></b>
            </div>
            <div class="col-xs-6 col-sm-6 col-md-6 col-lg-6">
                 Prepared by: <b><?php echo $prepared; ?></b>
            </div>
         </div>
         <div class="row">
            <div class="col-xs-6 col-sm-6 col-md-6 col-lg-6">
                 Address: <b><?php echo $address; ?></b>
            </div>
            <div class="col-xs-6 col-sm-6 col-md-6 col-lg-6">
                 Date: <b><?php echo $dispatchcontrolnewdate; ?></b>
            </div>
         </div>
         <div class="row">
            <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
                <center><h4><b>DISPATCH CONTROL</b></h4></center>
            </div>
         </div>
         <table class="table table-bordered">
            <thead class="table-header">
                <tr>
                    <th><center>CYCLE CODE</center></th>
                    <th><center>PICK-UP DATE</center></th>
                    <th><center>PRODUCT DESCRIPTION</center></th>
                    <th><center>UNIT</center></th>
                    <th><center>QUANTITY</center></th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td><center><?php echo $cyclecode1; ?></center></center></td>
                    <td><center><?php if($pickupdate1 != "") echo date("M d, Y", strtotime($pickupdate1)); ?></center></td>
                    <td><center><?php echo $sender1.' '.$deltype1; ?></center></td>
                    <td><center><?php if($cyclecode1 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty1 != "") { echo $qty1; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode2; ?></center></td>
                    <td><center><?php if($pickupdate2 != "") echo date("M d, Y", strtotime($pickupdate2)); ?></center></td>
                    <td><center><?php echo $sender2.' '.$deltype2; ?></center></td>
                     <td><center><?php if($cyclecode2 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty2 != "") { echo $qty2; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode3; ?></center></td>
                    <td><center><?php if($pickupdate3 != "") echo date("M d, Y", strtotime($pickupdate3)); ?></center></td>
                    <td><center><?php echo $sender3.' '.$deltype3; ?></center></td>
                    <td><center><?php if($cyclecode3 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty3 != "") { echo $qty3; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode4; ?></center></td>
                    <td><center><?php if($pickupdate4 != "") echo date("M d, Y", strtotime($pickupdate4)); ?></center></td>
                    <td><center><?php echo $sender4.' '.$deltype4; ?></center></td>
                    <td><center><?php if($cyclecode4 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty4 != "") { echo $qty4; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode5; ?></center></td>
                    <td><center><?php if($pickupdate5 != "") echo date("M d, Y", strtotime($pickupdate5)); ?></center></td>
                    <td><center><?php echo $sender5.' '.$deltype5; ?></td>
                    <td><center><?php if($cyclecode5 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty5 != "") { echo $qty5; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode6; ?></center></td>
                    <td><center><?php if($pickupdate6 != "") echo date("M d, Y", strtotime($pickupdate6)); ?></center></td>
                    <td><center><?php echo $sender6.' '.$deltype6; ?></center></td>
                    <td><center><?php if($cyclecode6 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty6 != "") { echo $qty6; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode7; ?></center></td>
                    <td><center><?php if($pickupdate7 != "") echo date("M d, Y", strtotime($pickupdate7)); ?></center></td>
                    <td><center><?php echo $sender7.' '.$deltype7; ?></center></td>
                    <td><center><?php if($cyclecode7 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty7 != "") { echo $qty7; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode8; ?></center></td>
                    <td><center><?php if($pickupdate8 != "") echo date("M d, Y", strtotime($pickupdate8)); ?></center></td>
                    <td><center><?php echo $sender8.' '.$deltype8; ?></center></td>
                    <td><center><?php if($cyclecode8 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty8 != "") { echo $qty8; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode9; ?></center></td>
                    <td><center><?php if($pickupdate9 != "") echo date("M d, Y", strtotime($pickupdate9)); ?></center></td>
                    <td><center><?php echo $sender9.' '.$deltype9; ?></center></td>
                    <td><center><?php if($cyclecode9 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty9 != "") { echo $qty9; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode10; ?></center></td>
                    <td><center><?php if($pickupdate10 != "") echo date("M d, Y", strtotime($pickupdate10)); ?></center></td>
                    <td><center><?php echo $sender10.' '.$deltype10; ?></center></td>
                    <td><center><?php if($cyclecode10 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty10 != "") { echo $qty10; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode11; ?></center></td>
                    <td><center><?php if($pickupdate11 != "") echo date("M d, Y", strtotime($pickupdate11)); ?></center></td>
                    <td><center><?php echo $sender11.' '.$deltype11; ?></center></td>
                    <td><center><?php if($cyclecode11 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty11 != "") { echo $qty11; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode12; ?></center></td>
                    <td><center><?php if($pickupdate12 != "") echo date("M d, Y", strtotime($pickupdate12)); ?></center></td>
                    <td><center><?php echo $sender12.' '.$deltype12; ?></center></td>
                    <td><center><?php if($cyclecode12 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty12 != "") { echo $qty12; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode13; ?></center></td>
                    <td><center><?php if($pickupdate13 != "") echo date("M d, Y", strtotime($pickupdate13)); ?></center></td>
                    <td><center><?php echo $sender13.' '.$deltype13; ?></center></td>
                    <td><center><?php if($cyclecode13 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty13 != "") { echo $qty13; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode14; ?></center></td>
                    <td><center><?php if($pickupdate14 != "") echo date("M d, Y", strtotime($pickupdate14)); ?></center></td>
                    <td><center><?php echo $sender14.' '.$deltype14; ?></center></td>
                    <td><center><?php if($cyclecode14 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty14 != "") { echo $qty14; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode15; ?></center></td>
                    <td><center><?php if($pickupdate15 != "") echo date("M d, Y", strtotime($pickupdate15)); ?></center></td>
                    <td><center><?php echo $sender15.' '.$deltype15; ?></center></td>
                    <td><center><?php if($cyclecode15 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty15 != "") { echo $qty15; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode16; ?></center></td>
                    <td><center><?php if($pickupdate16 != "") echo date("M d, Y", strtotime($pickupdate16)); ?></center></td>
                    <td><center><?php echo $sender16.' '.$deltype16; ?></center></td>
                    <td><center><?php if($cyclecode16 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty16 != "") { echo $qty16; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode17; ?></center></td>
                    <td><center><?php if($pickupdate17 != "") echo date("M d, Y", strtotime($pickupdate17)); ?></center></td>
                    <td><center><?php echo $sender17.' '.$deltype17; ?></center></td>
                    <td><center><?php if($cyclecode17 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty17 != "") { echo $qty17; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode18; ?></center></td>
                    <td><center><?php if($pickupdate18 != "") echo date("M d, Y", strtotime($pickupdate18)); ?></center></td>
                    <td><center><?php echo $sender18.' '.$deltype18; ?></center></td>
                    <td><center><?php if($cyclecode18 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty18 != "") { echo $qty18; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode19; ?></center></td>
                    <td><center><?php if($pickupdate19 != "") echo date("M d, Y", strtotime($pickupdate19)); ?></center></td>
                    <td><center><?php echo $sender19.' '.$deltype19; ?></center></td>
                    <td><center><?php if($cyclecode19 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty19 != "") { echo $qty19; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode20; ?></center></td>
                    <td><center><?php if($pickupdate20 != "") echo date("M d, Y", strtotime($pickupdate20)); ?></center></td>
                    <td><center><?php echo $sender20.' '.$deltype20; ?></center></td>
                    <td><center><?php if($cyclecode20 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty20 != "") { echo $qty20; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode21; ?></center></td>
                    <td><center><?php if($pickupdate21 != "") echo date("M d, Y", strtotime($pickupdate21)); ?></center></td>
                    <td><center><?php echo $sender21.' '.$deltype21; ?></td>
                    <td><center><?php if($cyclecode21 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty21 != "") { echo $qty21; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode22; ?></center></td>
                    <td><center><?php if($pickupdate22 != "") echo date("M d, Y", strtotime($pickupdate22)); ?></center></td>
                    <td><center><?php echo $sender22.' '.$deltype22; ?></td>
                    <td><center><?php if($cyclecode22 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty22 != "") { echo $qty22; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode23; ?></center></td>
                    <td><center><?php if($pickupdate23 != "") echo date("M d, Y", strtotime($pickupdate23)); ?></center></td>
                    <td><center><?php echo $sender23.' '.$deltype23; ?></td>
                    <td><center><?php if($cyclecode23 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty23 != "") { echo $qty23; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td><center><?php echo $cyclecode24; ?></center></td>
                    <td><center><?php if($pickupdate24 != "") echo date("M d, Y", strtotime($pickupdate24)); ?></center></td>
                    <td><center><?php echo $sender24.' '.$deltype24; ?></td>
                    <td><center><?php if($cyclecode24 != "") echo "Pieces"; ?></center></td>
                    <td><center><?php  if($qty24 != "") { echo $qty24; } else { echo "<span style='visibility: hidden;'>hello</span>"; } ?></center></td>
                </tr>
                <tr>
                    <td style="visibility: hidden;">hello</td>
                    <td style="visibility: hidden;">hello</td>
                    <td style="visibility: hidden;">hello</td>
                    <td><center><b>Total</b></center></td>
                    <td><center><?php  if($grandTotal != "") { echo $grandTotal; } else { echo "0"; } ?></center></td>
                </tr>
            </tbody>
         </table>
         <div class="row">
            <div class="col-xs-7 col-sm-7 col-md-7 col-lg-7">
                <div class="panel panel-default">
                  <div class="panel-heading"><center><b style="color: #ffffff !important;">IMPORTANT REMINDER</b></center></div>
                  <div class="panel-body"><center>POD's (Proof of Delivery) must return on/before the given timeline (NCR 8 days / PROVINCIAL 12 days after your pick-up). If not, the late POD's will not be compensated.</center></div>
                </div>
                <p><center>Please read the important reminder. Upon signing, it implies that you agreed the terms and condition</center></p>
            </div>
            <div class="col-xs-5 col-sm-5 col-md-5 col-lg-5">
              <img src="img/ppbsignature.png" style="position: absolute; width: 200px; margin-left: 40px; margin-top: 85px;" />
                <p>Received by:<br><br> ____________________________________<center style="font-size: 13px; text-align: center;">SIGNATURE OVER PRINTED NAME/DATE</center></p><br>
                <p>Noted:<br><center style="font-size: 16px; font-weight: bold;">PASTOR B. BENAVIDES</center><center style="font-size: 13px; padding: 0px;">PRESIDENT</center></p>
            </div>
         </div>
     </div>  

    <!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
    <!-- Include all compiled plugins (below), or include individual files as needed -->
    <script src="js/bootstrap.min.js"></script>
    <script type="text/javascript">
        window.onload = function() { window.print(); }
    </script>
  </body>
</html>