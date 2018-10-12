<?php
ini_set('max_execution_time', 19200); //300 sedb_connds = 5 minutes
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
if(isset($_GET["encodeid"])){
  $encodeid = $_GET['encodeid'];
}
?>
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- The above 3 meta tags *must* come first in the head; any other head content must come *after* these tags -->
    <title>PPBMS - Receipt</title>
    <!-- FavIcon -->
    <link rel="icon" type="image/png" href="img/fav.png" />
    <!-- Bootstrap -->
    <link href="css/bootstrap.min.css" rel="stylesheet">
    <script type="text/javascript">
        window.onload = function() { window.print(); }
    </script>

    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.3/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->
    <style type="text/css" media="print">
      @page { size: landscape; }
    </style>
    <style type="text/css">
        body {
           font-size: 8px;
        }
        .table-border {
            border: 1px solid #999;
        }
        .span-border {
            border: 1px solid #999;
            border-radius: 5px;
            width: 100%;
            height: 57px;
            padding: 4px 5px;
            margin: 10px 0px 10px 0px;
        }
        .rts-border {
            border: 1px solid #999;
            width: 100%;
            padding: 3px;
            margin: 10px 0px 20px 0px;
        }
        .span-empty {
            width: 100%;
            margin: 15px 0px 0px 0px;
        }
        .table-data {
            width: 100%;
        }
        .row-margin {
            padding: 0px 15px 0px 15px;
        }
        .table-header {
            visibility: hidden
        }
        .fillup-margin {
            padding: 5px 0px 0px 0px;
        }
        tr.highlight td {padding-left: 16px; padding-right:16px;}
        #page-break {page-break-after: always;}
        @page {
            size:  auto; 
            margin: 1mm;
        }
        @media print {
            #page-break {margin-top: 4mm;}
        }
        .particular {
            font-size: 11px;
        }
        .particular2 {
            font-size: 11px;
        }
        .cut-text { 
            font-family: Arial;
            width: 170px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            color: #252525;
            font-weight: bold;
        }
        #subsname {
            font-weight: bold;
            font-size: 11px;
        }
    </style>
  </head>
  <body>

     <div class="container-fluid">

            <?php
                $sql = "SELECT id, record_cycle_code, record_sender, record_subs_name, record_bar_code, record_area, record_address, record_deltype, record_pud, record_job_num, record_seq_num FROM ppbms_record WHERE encode_lists_id = '$encodeid' AND record_status_status = '1'";
                $query = mysqli_query($db_conn, $sql);
                $count = 0;
                while($row = mysqli_fetch_array($query)) {  
                    $count++; 
                    $recid = $row["id"];
                    $cyclecode = $row["record_cycle_code"];
                    $company = $row["record_sender"];
                    $subsname = $row["record_subs_name"];
                    $barcode = $row["record_bar_code"];
                    $area = $row["record_area"];
                    $address = $row["record_address"];
                    $deltype = $row["record_deltype"];
                    $pud = $row["record_pud"];
                    $jobnumber = $row["record_job_num"];
                    $seqnumber = $row["record_seq_num"];
                    $newPud = date("m/d/Y", strtotime($pud));
                    if($count % 6 != 0) {
                        if($count % 6 == 1) {
                    echo '
                    <div class="row" id="page-break">
                            <div class="col-xs-6 col-sm-6 col-md-6 col-lg-6 table-border">
                            <div class="row">
                            <table data-toggle="table" class="table-data">
                                <thead class="table-header">
                                <tr>
                                    <th>Column 1</th>
                                    <th>Column 2</th>
                                    <th>Column 1</th>
                                    <th>Column 2</th>
                                </tr>
                                </thead>
                                <tbody>
                                <tr class="highlight">
                                    <td class="particular2" colspan="2">
                                        PARTICULAR ><br>
                                        <div class="cut-text">Sender: '.$company.'</div>
                                        <div class="cut-text">DelType: '.$deltype.'</div>
                                        <div class="cut-text">Area: '.$area.'</div>
                                    </td>
                                    <td class="particular2 pudcc" colspan="2">
                                        <div class="cut-text">Cycle: '.$cyclecode.'</div>
                                        <div class="cut-text">PUD: '.$pud.'</div>
                                        <div class="cut-text">Job Order #: '.$jobnumber.'</div>
                                        <div class="cut-text">POD #: '.$count.'</div>
                                    </td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="4">
                                        <div class="span-border">'.$barcode.'<br><span id="subsname">'.$subsname.'</span><br>'.$address.'</div>
                                    </td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="4"><span style="font-size: 6px;"><b>I hereby acknowledge receipt of the particulars details below</b></span></td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="4">
                                        <div class="span-empty"></div>
                                    </td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="2">
                                        ________________________________<br><span style="font-size: 6px;"><b>RECEIVED BY(NAME WITH SIGNATURE)</b></span>
                                    </td>
                                    <td colspan="2">
                                        ____________________________________<br><span style="font-size: 6px;"><b>RELATION TO ADDRESSEE</b></span>
                                    </td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="2" class="fillup-margin">
                                        ________________________________<br><span style="font-size: 6px;"><b>DELIVERED BY</b></span>
                                    </td>
                                    <td colspan="2" class="fillup-margin">
                                        ____________________________________<br><span style="font-size: 6px;"><b>DATE RECEIVED</b></span>
                                    </td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="2">
                                    <div class="rts-border">
                                        <b>RTS REMARKS</b><br>
                                        1 [ ] Moved Out (Person / Company)<br>
                                        2 [ ] Unknown (Person / Company)<br>
                                        3 [ ] Incomplete / Incorrect Address<br>
                                        4 [ ] Refused To Accept<br>
                                        5 [ ] House Closed / No One To Receive<br>
                                        6 [ ] Others __________________
                                    </div>
                                    </td>
                                    <td class="particular" colspan="2">
                                        <img alt="Coding Sips" src="includes/barcode.php?text='.$barcode.'&orientation=horizontal&size=40" style="width: 100%; height: 25px; margin-left: -5px;" />
                                        <b>PPB Messengerial Services, Inc.</b><br>
                                        <span>Tel. No. 478-2822</span>
                                    </td>
                                </tr>
                                </tbody>
                            </table>
                                </div>
                            </div>
                            
                    ';
                    } else {
                        echo '
                        <div class="col-xs-6 col-sm-6 col-md-6 col-lg-6 table-border">
                            <div class="row">
                            <table data-toggle="table" class="table-data">
                                <thead class="table-header">
                                <tr>
                                    <th>Column 1</th>
                                    <th>Column 2</th>
                                    <th>Column 1</th>
                                    <th>Column 2</th>
                                </tr>
                                </thead>
                                <tbody>
                                    <tr class="highlight">
                                    <td class="particular2" colspan="2">
                                        PARTICULAR ><br>
                                        <div class="cut-text">Sender: '.$company.'</div>
                                        <div class="cut-text">DelType: '.$deltype.'</div>
                                        <div class="cut-text">Area: '.$area.'</div>
                                    </td>
                                    <td class="particular2 pudcc" colspan="2">
                                        <div class="cut-text">Cycle: '.$cyclecode.'</div>
                                        <div class="cut-text">PUD: '.$pud.'</div>
                                        <div class="cut-text">Job Order #: '.$jobnumber.'</div>
                                        <div class="cut-text">POD #: '.$count.'</div>
                                    </td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="4">
                                        <div class="span-border">'.$barcode.'<br><span id="subsname">'.$subsname.'</span><br>'.$address.'</div>
                                    </td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="4"><span style="font-size: 6px;"><b>I hereby acknowledge receipt of the particulars details below</b></span></td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="4">
                                        <div class="span-empty"></div>
                                    </td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="2">
                                        ________________________________<br><span style="font-size: 6px;"><b>RECEIVED BY(NAME WITH SIGNATURE)</b></span>
                                    </td>
                                    <td colspan="2">
                                        ____________________________________<br><span style="font-size: 6px;"><b>RELATION TO ADDRESSEE</b></span>
                                    </td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="2" class="fillup-margin">
                                        ________________________________<br><span style="font-size: 6px;"><b>DELIVERED BY</b></span>
                                    </td>
                                    <td colspan="2" class="fillup-margin">
                                        ____________________________________<br><span style="font-size: 6px;"><b>DATE RECEIVED</b></span>
                                    </td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="2">
                                    <div class="rts-border">
                                        <b>RTS REMARKS</b><br>
                                        1 [ ] Moved Out (Person / Company)<br>
                                        2 [ ] Unknown (Person / Company)<br>
                                        3 [ ] Incomplete / Incorrect Address<br>
                                        4 [ ] Refused To Accept<br>
                                        5 [ ] House Closed / No One To Receive<br>
                                        6 [ ] Others __________________
                                    </div>
                                    </td>
                                    <td class="particular" colspan="2">
                                        <img alt="Coding Sips" src="includes/barcode.php?text='.$barcode.'&orientation=horizontal&size=40" style="width: 100%; height: 25px; margin-left: -5px;" />
                                        <b>PPB Messengerial Services, Inc.</b><br>
                                        <span>Tel. No. 478-2822</span>
                                    </td>
                                </tr>
                                </tbody>
                            </table>
                                </div>
                            </div>
                        ';
                    }
                } else {
                    echo '
                    <div class="col-xs-6 col-sm-6 col-md-6 col-lg-6 table-border">
                            <div class="row">
                            <table data-toggle="table" class="table-data">
                                <thead class="table-header">
                                <tr>
                                    <th>Column 1</th>
                                    <th>Column 2</th>
                                    <th>Column 1</th>
                                    <th>Column 2</th>
                                </tr>
                                </thead>
                                <tbody>
                                    <tr class="highlight">
                                    <td class="particular2" colspan="2">
                                        PARTICULAR ><br>
                                        <div class="cut-text">Sender: '.$company.'</div>
                                        <div class="cut-text">DelType: '.$deltype.'</div>
                                        <div class="cut-text">Area: '.$area.'</div>
                                    </td>
                                    <td class="particular2 pudcc" colspan="2">
                                        <div class="cut-text">Cycle: '.$cyclecode.'</div>
                                        <div class="cut-text">PUD: '.$pud.'</div>
                                        <div class="cut-text">Job Order #: '.$jobnumber.'</div>
                                        <div class="cut-text">POD #: '.$count.'</div>
                                    </td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="4">
                                        <div class="span-border">'.$barcode.'<br><span id="subsname">'.$subsname.'</span><br>'.$address.'</div>
                                    </td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="4"><span style="font-size: 6px;"><b>I hereby acknowledge receipt of the particulars details below</b></span></td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="4">
                                        <div class="span-empty"></div>
                                    </td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="2">
                                        ________________________________<br><span style="font-size: 6px;"><b>RECEIVED BY(NAME WITH SIGNATURE)</b></span>
                                    </td>
                                    <td colspan="2">
                                        ____________________________________<br><span style="font-size: 6px;"><b>RELATION TO ADDRESSEE</b></span>
                                    </td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="2" class="fillup-margin">
                                        ________________________________<br><span style="font-size: 6px;"><b>DELIVERED BY</b></span>
                                    </td>
                                    <td colspan="2" class="fillup-margin">
                                        ____________________________________<br><span style="font-size: 6px;"><b>DATE RECEIVED</b></span>
                                    </td>
                                </tr>
                                <tr class="highlight">
                                    <td colspan="2">
                                    <div class="rts-border">
                                        <b>RTS REMARKS</b><br>
                                        1 [ ] Moved Out (Person / Company)<br>
                                        2 [ ] Unknown (Person / Company)<br>
                                        3 [ ] Incomplete / Incorrect Address<br>
                                        4 [ ] Refused To Accept<br>
                                        5 [ ] House Closed / No One To Receive<br>
                                        6 [ ] Others __________________
                                    </div>
                                    </td>
                                    <td class="particular" colspan="2">
                                        <img alt="Coding Sips" src="includes/barcode.php?text='.$barcode.'&orientation=horizontal&size=40" style="width: 100%; height: 25px; margin-left: -5px;" />
                                        <b>PPB Messengerial Services, Inc.</b><br>
                                        <span>Tel. No. 478-2822</span>
                                    </td>
                                </tr>
                                </tbody>
                            </table>
                                </div>
                            </div>
              </div>
                    ';
                  
                }
                }
                mysqli_close($db_conn);
            ?>


    </div>

    <!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
    <!-- Include all compiled plugins (below), or include individual files as needed -->
    <script src="js/bootstrap.min.js"></script>
  </body>
</html>