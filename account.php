<?php
ini_set('max_execution_time', 19200); //300 sedb_connds = 5 minutes
include_once("includes/loginstatus.php");
if (!isset($_SESSION["userid"])) {
  header("location: index.php");

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
$classactivedispatchcontrol = "";
$classactiveuploadbackup = "";
$classactivereports = "";
$echo = "";
$messengerid = "";
$encodeid = "0";

// GET USER ID
if(isset($_GET["id"])){
  $id = preg_replace('#[^a-z0-9-]#i', '', $_GET['id']);
} else {
  header("location: index.php");
 
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

$sql_imported_lists = "SELECT COUNT(id) AS count_num FROM ppbms_encode_list WHERE encode_lists_status = '1'";
$query_imported_lists = mysqli_query($db_conn, $sql_imported_lists);
while($row1 = mysqli_fetch_array($query_imported_lists)) {  
	$count_imported_lists = $row1["count_num"];
}

$sql_record = "SELECT COUNT(id) AS count_num FROM ppbms_record WHERE record_status_status = '1'";
$query_record = mysqli_query($db_conn, $sql_record);
while($row2 = mysqli_fetch_array($query_record)) {  
	$count_record = $row2["count_num"];
}

$sql_dispatch_control_messenger = "SELECT COUNT(id) AS count_num FROM ppbms_dispatch_control_messenger WHERE dispatch_control_messenger_status = '1'";
$query_dispatch_control_messenger = mysqli_query($db_conn, $sql_dispatch_control_messenger);
while($row3 = mysqli_fetch_array($query_dispatch_control_messenger)) {  
	$count_dispatch_control_messenger = $row3["count_num"];
}

$sql_lists_completed = "SELECT COUNT(id) AS count_num FROM ppbms_record WHERE record_date_receive != '' AND record_receive_by != '' AND record_relation != '' AND record_messenger != '' AND record_status != '' AND record_status_status = '1'";
$query_lists_completed = mysqli_query($db_conn, $sql_lists_completed);
while($row4 = mysqli_fetch_array($query_lists_completed)) {  
	$count_lists_completed = $row4["count_num"];
}

// ROUTE
if(

  (
    !isset($_GET["dashboard"]) 
    && !isset($_GET["settings"])
    && !isset($_GET["masterlists"])
    && !isset($_GET["dispatchcontrol"])
    && !isset($_GET["encodemasterlists"])
    && !isset($_GET["importmasterlists"])
    && !isset($_GET["reports"])
    && !isset($_GET["reportsettings"])            
    )

// IF ADMIN LEVEL
  || (isset($_GET["dashboard"]) && trim($_GET["dashboard"]) != "focus") 
  
  || (isset($_GET["settings"]) && trim($_GET["settings"]) != "focus") 

  || (isset($_GET["masterlists"]) && trim($_GET["masterlists"]) != "focus")

  || (isset($_GET["dispatchcontrol"]) && trim($_GET["dispatchcontrol"]) != "focus")  

  || (isset($_GET["encodemasterlists"]) && trim($_GET["encodemasterlists"]) != "focus") 

  || (isset($_GET["importmasterlists"]) && trim($_GET["importmasterlists"]) != "focus")

  || (isset($_GET["reports"]) && trim($_GET["reports"]) != "focus")

  || (isset($_GET["reportsettings"]) && trim($_GET["reportsettings"]) != "focus")    

  ) {
  $pheading = "Erorr 404";
  $pbody = "Page not found.";
} else if (isset($_GET["dashboard"]) && trim($_GET["dashboard"]) == "focus"){
  $pheading = "Dashboard";
  $classactivedashboard = "active";
} else if (isset($_GET["settings"]) && trim($_GET["settings"]) == "focus"){
  $pheading = "Account Settings";
  $classactivesettings = "active";
} else if (isset($_GET["masterlists"]) && isset($_GET["select"]) && trim($_GET["masterlists"]) == "focus" && trim($_GET["select"]) == "add"){
  $pheading = "Import Master Lists";
  $classactivemasterlists = "active";
} else if (isset($_GET["masterlists"]) && isset($_GET["select"]) && trim($_GET["masterlists"]) == "focus" && trim($_GET["select"]) == "view"){
  $pheading = "View Master Lists";
  $classactivemasterlists = "active";
} else if (isset($_GET["masterlists"]) && isset($_GET["select"]) && trim($_GET["masterlists"]) == "focus" && trim($_GET["select"]) == "barcode"){
  $pheading = "Master Lists Settings";
  $classactivemasterlists = "active";
} else if (isset($_GET["dispatchcontrol"]) && isset($_GET["select"]) && trim($_GET["dispatchcontrol"]) == "focus" && trim($_GET["select"]) == "add"){
  $pheading = "Add Dispatch Control";
  $classactivedispatchcontrol = "active";
} else if (isset($_GET["dispatchcontrol"]) && isset($_GET["select"]) && trim($_GET["dispatchcontrol"]) == "focus" && trim($_GET["select"]) == "view"){
  $pheading = "View Dispatch Control";
  $classactivedispatchcontrol = "active";
} else if (isset($_GET["encodemasterlists"]) && trim($_GET["encodemasterlists"]) == "focus"){
  $pheading = "Manual Encode Master Lists";
  $classactiveencodemasterlists = "active";
} else if (isset($_GET["importmasterlists"]) && trim($_GET["importmasterlists"]) == "focus"){
  $pheading = "Import to Encode Master Lists";
  $classactiveencodemasterlists = "active";
} else if (isset($_GET["reports"]) && trim($_GET["reports"]) == "focus"){
  $pheading = "Export Report";
  $classactivereports = "active";
} else if (isset($_GET["reportsettings"]) && trim($_GET["reportsettings"]) == "focus"){
  $pheading = "Report Settings";
  $classactivereports = "active";
}  else {
  $pheading = "Erorr 404";
  $pbody = "Page not found.";
}

?>

<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PPBMS - <?php echo $pheading; ?></title>

<link href="css/bootstrap.min.css" rel="stylesheet">
<!-- <link href="css/datepicker3.css" rel="stylesheet"> -->
<link href="css/styles.css" rel="stylesheet">
<link href="css/custom/account.css?refresh=true" rel="stylesheet">
<!-- FavIcon -->
<link rel="icon" type="image/png" href="img/fav.png" />
<!--Icons-->
<script src="js/lumino.glyphs.js"></script>
<!-- Google Fonts -->
<link href="https://fonts.googleapis.com/css?family=Roboto" rel="stylesheet">
<link rel="stylesheet" href="css/font-awesome.min.css">
<!-- <link rel="stylesheet" href="//cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.3.0/css/datepicker.min.css" />
<link rel="stylesheet" href="//cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.3.0/css/datepicker3.min.css" /> -->
<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
	<script src="js/main.js"></script>
    <script src="js/ajax.js"></script>
<script>
  $( function() {
  	$( "#add-dispatch-control-messenger-date" ).datepicker();
  	$( "#add-dispatch-control-data-pickup-date" ).datepicker();
  	$( "#update-record-date-recieve" ).datepicker();
    $( "#update-record-date-reported" ).datepicker();
    $( "#edit-pud-date" ).datepicker();
    $( "#edit-date-receive-date" ).datepicker();
    $( "#edit-date-reported-date" ).datepicker();
    $( "#search-encode-lists-date" ).datepicker();
  } );
  function request_page_on_load(pn){
	    	var encodeid = "<?php echo $encodeid; ?>";
	    	var txtSearch = "";
	           var txtFilter = "all";    
			 var results_box = _("record-search-result");

			 $.ajax({  
	                url:"parsers/fetch-search-count.php",  
	                method:"post",  
	                data:{search:txtSearch, filter:txtFilter, encodeid:encodeid},  
	                dataType:"text",  
	                success: function(data) {

			 var rpp = 10; // results per page
			 var last = data; // last page number
				var pagination_controls1 = _("pagination_controls1");
				var pagination_controls2 = _("pagination_controls2");
				results_box.innerHTML = "loading results ...";
			    var ajax = ajaxObj("POST", "parsers/fetch-master-lists.php");
	            ajax.onreadystatechange = function() {
	              if(ajax.readyState == 4 && ajax.status == 200) {
						var dataArray = ajax.responseText.split("||");
						var html_output = "";
						html_output += '<table class="table table-striped"><thead><tr><th>REC#</th><th>SENDER</th><th>DELTYPE</th><th>PUD</th><th>Month</th><th>JOB#</th><th>Checklist For PPB</th><th>File Name</th><th>seq no.</th><th>CYCLECODE</th><th>Qty</th><th>ADDRESS</th><th>AREA</th><th>SUBSCRIBERS NAME</th><th>BARCODE</th><th>ACCT NO</th><th>DATE RECEIVED</th><th>RECEIVED BY</th><th>RELATION</th><th>MESSENGER</th><th>STATUS</th><th>Reason for RTS</th><th>Remarks</th><th>Date Reported</th><th>Action</th></tr></thead><tbody>';
					    for(i = 0; i < dataArray.length - 1; i++){
							var itemArray = dataArray[i].split("|");
							html_output += '<tr><td>'+itemArray[0]+'</td><td>'+itemArray[1]+'</td><td>'+itemArray[2]+'</td><td>'+itemArray[3]+'</td><td>'+itemArray[4]+'</td><td>'+itemArray[5]+'</td><td>'+itemArray[6]+'</td><td>'+itemArray[7]+'</td><td>'+itemArray[8]+'</td><td>'+itemArray[9]+'</td><td>'+itemArray[10]+'</td><td>'+itemArray[11]+'</td><td>'+itemArray[12]+'</td><td>'+itemArray[13]+'</td><td>'+itemArray[14]+'</td><td>'+itemArray[15]+'</td><td>'+itemArray[16]+'</td><td>'+itemArray[17]+'</td><td>'+itemArray[18]+'</td><td>'+itemArray[19]+'</td><td>'+itemArray[20]+'</td><td>'+itemArray[21]+'</td><td>'+itemArray[22]+'</td><td>'+itemArray[23]+'</td><td><a href="account.php?id=1&masterlists=focus&select=view&encodeid='+encodeid+'&editrecid='+itemArray[0]+'">Edit</a> | <a href="#"  onclick="openMasterListItemDeleteDialog(\''+itemArray[0]+'\')">Delete</a></td></td></tr>';
						}
						html_output += '</tbody></table>';
						results_box.innerHTML = html_output;
						console.log(ajax.responseText);
				    }
	            }
			    ajax.send("rpp="+rpp+"&last="+last+"&pn="+pn+"&search="+txtSearch+"&filter="+txtFilter+"&encodeid="+encodeid);
				// Change the pagination controls
				var paginationCtrls = "";
			    // Only if there is more than 1 page worth of results give the user pagination controls
			    if(last != 1){
					if (pn > 1) {
						paginationCtrls += '<button class="btn btn-default" id="minusBtn" onclick="request_page('+(pn-1)+')">&lt;</button>';
			    	}
					paginationCtrls += ' &nbsp; &nbsp; <b>Page '+pn+' of '+last+'</b> &nbsp; &nbsp; ';
			    	if (pn != last) {
			        	paginationCtrls += '<button class="btn btn-default" id="addBtn" onclick="request_page('+(pn+1)+')">&gt;</button>';
			    	}
			    }
				pagination_controls1.innerHTML = paginationCtrls;
				pagination_controls2.innerHTML = paginationCtrls;

				}

	            }); 
		}
  function request_page_not_recieve(pn, cyclecode, messengerid){  
			 var results_box = _("record-search-result-not-recieve");

			 $.ajax({  
	                url:"parsers/fetch-count-not-recieve.php",  
	                method:"post",  
	                data:{cyclecode:cyclecode, messengerid:messengerid},  
	                dataType:"text",  
	                success: function(data) {

			 var rpp = 10; // results per page
			 var last = data; // last page number
				var pagination_controls1 = _("pagination_controls_not_recieve_1");
				var pagination_controls2 = _("pagination_controls_not_recieve_2");
				results_box.innerHTML = "loading results ...";
			    var ajax = ajaxObj("POST", "parsers/fetch-master-lists-not-recieve.php");
	            ajax.onreadystatechange = function() {
	              if(ajax.readyState == 4 && ajax.status == 200) {
						var dataArray = ajax.responseText.split("||");
						var html_output = "";
						html_output += '<table class="table table-striped"><thead><tr><th>REC#</th><th>SENDER</th><th>DELTYPE</th><th>PUD</th><th>Month</th><th>JOB#</th><th>Checklist For PPB</th><th>File Name</th><th>seq no.</th><th>CYCLECODE</th><th>Qty</th><th>ADDRESS</th><th>AREA</th><th>SUBSCRIBERS NAME</th><th>BARCODE</th><th>ACCT NO</th><th>DATE RECEIVED</th><th>RECEIVED BY</th><th>RELATION</th><th>MESSENGER</th><th>STATUS</th><th>Reason for RTS</th><th>Remarks</th><th>Date Reported</th><th>Action</th></tr></thead><tbody>';
					    for(i = 0; i < dataArray.length - 1; i++){
							var itemArray = dataArray[i].split("|");
							html_output += '<tr><td>'+itemArray[0]+'</td><td>'+itemArray[1]+'</td><td>'+itemArray[2]+'</td><td>'+itemArray[3]+'</td><td>'+itemArray[4]+'</td><td>'+itemArray[5]+'</td><td>'+itemArray[6]+'</td><td>'+itemArray[7]+'</td><td>'+itemArray[8]+'</td><td>'+itemArray[9]+'</td><td>'+itemArray[10]+'</td><td>'+itemArray[11]+'</td><td>'+itemArray[12]+'</td><td>'+itemArray[13]+'</td><td>'+itemArray[14]+'</td><td>'+itemArray[15]+'</td><td>'+itemArray[16]+'</td><td>'+itemArray[17]+'</td><td>'+itemArray[18]+'</td><td>'+itemArray[19]+'</td><td>'+itemArray[20]+'</td><td>'+itemArray[21]+'</td><td>'+itemArray[22]+'</td><td>'+itemArray[23]+'</td><td>Edit | <a href="#" onclick="openMasterListItemDeleteDialog(\''+itemArray[0]+'\')">Delete</a></td></td></tr>';
						}
						html_output += '</tbody></table>';
						results_box.innerHTML = html_output;
						console.log(ajax.responseText);
				    }
	            }
			    ajax.send("rpp="+rpp+"&last="+last+"&pn="+pn+"&cyclecode="+cyclecode+"&messengerid="+messengerid);
				// Change the pagination controls
				var paginationCtrls = "";
			    // Only if there is more than 1 page worth of results give the user pagination controls
			    if(last != 1){
					if (pn > 1) {
						paginationCtrls += '<button class="btn btn-default" id="minusBtn" onclick="request_page_not_recieve('+(pn-1)+','+cyclecode+','+messengerid+')">&lt;</button>';
			    	}
					paginationCtrls += ' &nbsp; &nbsp; <b>Page '+pn+' of '+last+'</b> &nbsp; &nbsp; ';
			    	if (pn != last) {
			        	paginationCtrls += '<button class="btn btn-default" id="addBtn" onclick="request_page_not_recieve('+(pn+1)+','+cyclecode+','+messengerid+')">&gt;</button>';
			    	}
			    }
				pagination_controls1.innerHTML = paginationCtrls;
				pagination_controls2.innerHTML = paginationCtrls;

				}

	            }); 
		}
	function request_page_not_assign(pn, encodeid){  
			 var results_box = _("record-search-result-not-assign");

			 $.ajax({  
	                url:"parsers/fetch-count-not-assign.php",  
	                method:"post",  
	                data:{encodeid:encodeid},  
	                dataType:"text",  
	                success: function(data) {

			 var rpp = 10; // results per page
			 var last = data; // last page number
				var pagination_controls1 = _("pagination_controls_not_assign_1");
				var pagination_controls2 = _("pagination_controls_not_assign_2");
				results_box.innerHTML = "loading results ...";
			    var ajax = ajaxObj("POST", "parsers/fetch-master-lists-not-assign.php");
	            ajax.onreadystatechange = function() {
	              if(ajax.readyState == 4 && ajax.status == 200) {
						var dataArray = ajax.responseText.split("||");
						var html_output = "";
						html_output += '<table class="table table-striped"><thead><tr><th>REC#</th><th>SENDER</th><th>DELTYPE</th><th>PUD</th><th>Month</th><th>JOB#</th><th>Checklist For PPB</th><th>File Name</th><th>seq no.</th><th>CYCLECODE</th><th>Qty</th><th>ADDRESS</th><th>AREA</th><th>SUBSCRIBERS NAME</th><th>BARCODE</th><th>ACCT NO</th><th>DATE RECEIVED</th><th>RECEIVED BY</th><th>RELATION</th><th>MESSENGER</th><th>STATUS</th><th>Reason for RTS</th><th>Remarks</th><th>Date Reported</th></tr></thead><tbody>';
					    for(i = 0; i < dataArray.length - 1; i++){
							var itemArray = dataArray[i].split("|");
							html_output += '<tr><td>'+itemArray[0]+'</td><td>'+itemArray[1]+'</td><td>'+itemArray[2]+'</td><td>'+itemArray[3]+'</td><td>'+itemArray[4]+'</td><td>'+itemArray[5]+'</td><td>'+itemArray[6]+'</td><td>'+itemArray[7]+'</td><td>'+itemArray[8]+'</td><td>'+itemArray[9]+'</td><td>'+itemArray[10]+'</td><td>'+itemArray[11]+'</td><td>'+itemArray[12]+'</td><td>'+itemArray[13]+'</td><td>'+itemArray[14]+'</td><td>'+itemArray[15]+'</td><td>'+itemArray[16]+'</td><td>'+itemArray[17]+'</td><td>'+itemArray[18]+'</td><td>'+itemArray[19]+'</td><td>'+itemArray[20]+'</td><td>'+itemArray[21]+'</td><td>'+itemArray[22]+'</td><td>'+itemArray[23]+'</td></tr>';
						}
						html_output += '</tbody></table>';
						results_box.innerHTML = html_output;
						console.log(ajax.responseText);
				    }
	            }
			    ajax.send("rpp="+rpp+"&last="+last+"&pn="+pn+"&encodeid="+encodeid);
				// Change the pagination controls
				var paginationCtrls = "";
			    // Only if there is more than 1 page worth of results give the user pagination controls
			    if(last != 1){
					if (pn > 1) {
						paginationCtrls += '<button class="btn btn-default" id="minusBtn" onclick="request_page_not_assign('+(pn-1)+','+encodeid+')">&lt;</button>';
			    	}
					paginationCtrls += ' &nbsp; &nbsp; <b>Page '+pn+' of '+last+'</b> &nbsp; &nbsp; ';
			    	if (pn != last) {
			        	paginationCtrls += '<button class="btn btn-default" id="addBtn" onclick="request_page_not_assign('+(pn+1)+','+encodeid+')">&gt;</button>';
			    	}
			    }
				pagination_controls1.innerHTML = paginationCtrls;
				pagination_controls2.innerHTML = paginationCtrls;

				}

	            }); 
		}

		function request_page_dispatch(pn){  
			 var results_box = _("record-search-result-dispatch");

			 $.ajax({  
	                url:"parsers/fetch-count-dispatch.php",  
	                method:"post",  
	                dataType:"text",  
	                success: function(data) {

			 var rpp = 10; // results per page
			 var last = data; // last page number
				var pagination_controls1 = _("pagination_controls_dispatch_1");
				var pagination_controls2 = _("pagination_controls_dispatch_2");
				results_box.innerHTML = "loading results ...";
			    var ajax = ajaxObj("POST", "parsers/fetch-dispatch.php");
	            ajax.onreadystatechange = function() {
	              if(ajax.readyState == 4 && ajax.status == 200) {
						var dataArray = ajax.responseText.split("||");
						var html_output = "";
						html_output += '<table class="table table-striped"><thead><tr><th>No.</th><th>Messenger Name</th><th>Address</th><th>Prepared</th><th>Date</th><th>Action</th></tr></thead><tbody>';
					    for(i = 0; i < dataArray.length - 1; i++){
							var itemArray = dataArray[i].split("|");
							html_output += '<tr><td>'+itemArray[0]+'</td><td>'+itemArray[1]+'</td><td>'+itemArray[2]+'</td><td>'+itemArray[3]+'</td><td>'+itemArray[4]+'</td><td><a href="account.php?id=<?php echo $log_id; ?>&dispatchcontrol=focus&select=view&recordidview='+itemArray[0]+'" id="'+itemArray[0]+'">View Record</a> | <a target="_blank" href="dispatchcontrol.php?id=<?php echo $log_id; ?>&messengerid='+itemArray[0]+'" id="'+itemArray[0]+'">Print Receipt</a> | <a target="_blank" href="dispatchproof.php?id=<?php echo $log_id; ?>&messengerid='+itemArray[0]+'" id="'+itemArray[0]+'">Print Proof</a> | <a href="#" id="'+itemArray[0]+'" onclick="openDispatchControlMessengerUpdateDialog(\''+itemArray[0]+'\',\''+itemArray[1]+'\',\''+itemArray[2]+'\',\''+itemArray[3]+'\',\''+itemArray[5]+'\')">Edit</a> | <a href="#" id="'+itemArray[0]+'" onclick="openDispatchControlMessengerDeleteDialog(\''+itemArray[0]+'\')">Delete</a></td></tr>';
						}
						html_output += '</tbody></table>';
						results_box.innerHTML = html_output;
						console.log(ajax.responseText);
				    }
	            }
			    ajax.send("rpp="+rpp+"&last="+last+"&pn="+pn);
				// Change the pagination controls
				var paginationCtrls = "";
			    // Only if there is more than 1 page worth of results give the user pagination controls
			    if(last != 1){
					if (pn > 1) {
						paginationCtrls += '<button class="btn btn-default" id="minusBtn" onclick="request_page_dispatch('+(pn-1)+')">&lt;</button>';
			    	}
					paginationCtrls += ' &nbsp; &nbsp; <b>Page '+pn+' of '+last+'</b> &nbsp; &nbsp; ';
			    	if (pn != last) {
			        	paginationCtrls += '<button class="btn btn-default" id="addBtn" onclick="request_page_dispatch('+(pn+1)+')">&gt;</button>';
			    	}
			    }
				pagination_controls1.innerHTML = paginationCtrls;
				pagination_controls2.innerHTML = paginationCtrls;

				}

	            }); 
		}

		function request_page_encode_lists(pn){  
			 var results_box = _("record-search-result-encode-lists");

			 $.ajax({  
	                url:"parsers/fetch-count-encode-lists.php",  
	                method:"post",  
	                dataType:"text",  
	                success: function(data) {

	                	console.log(data);

			 var rpp = 10; // results per page
			 var last = data; // last page number
				var pagination_controls1 = _("pagination_controls_encode_lists_1");
				var pagination_controls2 = _("pagination_controls_encode_lists_2");
				results_box.innerHTML = "loading results ...";
			    var ajax = ajaxObj("POST", "parsers/fetch-encode-lists.php");
	            ajax.onreadystatechange = function() {
	              if(ajax.readyState == 4 && ajax.status == 200) {
						var dataArray = ajax.responseText.split("||");
						var html_output = "";
						html_output += '<table class="table table-striped"><thead><tr><th>No.</th><th>File Name</th><th>Date Imported</th><th>Assigned Count</th><th>Unassigned Count</th><th>Total Count</th><th>Action</th></tr></thead><tbody>';
					    for(i = 0; i < dataArray.length - 1; i++){
							var itemArray = dataArray[i].split("|");
							html_output += '<tr><td>'+itemArray[0]+'</td><td>'+itemArray[1]+'</td><td>'+itemArray[2]+'</td><td>'+itemArray[3]+'</td><td><a id="'+itemArray[0]+'" href="account.php?id=<?php echo $log_id; ?>&masterlists=focus&select=view&encodeid='+itemArray[0]+'&check=unassigned">'+itemArray[4]+'</a></td><td>'+itemArray[5]+'</td><td><a id="'+itemArray[0]+'" href="account.php?id=<?php echo $log_id; ?>&masterlists=focus&select=view&encodeid='+itemArray[0]+'">View Record</a> | <a target="_blank" id="'+itemArray[0]+'" href="receipt.php?id=<?php echo $log_id; ?>&encodeid='+itemArray[0]+'">Print Receipt</a> | <a target="_blank" id="'+itemArray[0]+'" href="receiptold.php?id=<?php echo $log_id; ?>&encodeid='+itemArray[0]+'">Print Old Receipt</a> | <a href="#" id="'+itemArray[0]+'" onclick="openEncodeListsDeleteDialog(\''+itemArray[0]+'\')">Delete</a></td></tr>';
						}
						html_output += '</tbody></table>';
						results_box.innerHTML = html_output;
						console.log(ajax.responseText);
				    }
	            }
			    ajax.send("rpp="+rpp+"&last="+last+"&pn="+pn);
				// Change the pagination controls
				var paginationCtrls = "";
			    // Only if there is more than 1 page worth of results give the user pagination controls
			    if(last != 1){
					if (pn > 1) {
						paginationCtrls += '<button class="btn btn-default" id="minusBtn" onclick="request_page_encode_lists('+(pn-1)+')">&lt;</button>';
			    	}
					paginationCtrls += ' &nbsp; &nbsp; <b>Page '+pn+' of '+last+'</b> &nbsp; &nbsp; ';
			    	if (pn != last) {
			        	paginationCtrls += '<button class="btn btn-default" id="addBtn" onclick="request_page_encode_lists('+(pn+1)+')">&gt;</button>';
			    	}
			    }
				pagination_controls1.innerHTML = paginationCtrls;
				pagination_controls2.innerHTML = paginationCtrls;

				}

	            }); 
		}
</script>
<!--[if lt IE 9]>
<script src="js/html5shiv.js"></script>
<script src="js/respond.min.js"></script>
<![endif]-->
</head>

<body>
	<nav class="navbar navbar-inverse navbar-fixed-top" role="navigation">
		<div class="container-fluid">
			<div class="navbar-header">
				<button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#sidebar-collapse">
					<span class="sr-only">Toggle navigation</span>
					<span class="icon-bar"></span>
					<span class="icon-bar"></span>
					<span class="icon-bar"></span>
				</button>
				<a class="navbar-brand" href="#"><span>PPB</span>MS</a>
				<ul class="user-menu">
					<li class="dropdown pull-right">
						<a href="#" class="dropdown-toggle" data-toggle="dropdown"><svg class="glyph stroked male-user"><use xlink:href="#stroked-male-user"></use></svg> <?php echo $accname ?> <span class="caret"></span></a>
						<ul class="dropdown-menu" role="menu">
							<li><a href="logout.php"><svg class="glyph stroked cancel"><use xlink:href="#stroked-cancel"></use></svg> Logout</a></li>
						</ul>
					</li>
				</ul>
			</div>
							
		</div><!-- /.container-fluid -->
	</nav>
		
	<div id="sidebar-collapse" class="col-sm-3 col-lg-2 sidebar">
		<ul class="nav menu">
			<li class="<?php echo $classactivedashboard; ?>"><a href="account.php?id=<?php echo $id; ?>&dashboard=focus"><svg class="glyph stroked dashboard-dial"><use xlink:href="#stroked-dashboard-dial"></use></svg> Dashboard</a></li>
			<li class="parent <?php echo $classactivemasterlists ?>">
				<a data-toggle="collapse" href="#sub-item-1">
					<span data-toggle="collapse" href="#sub-item-1"><svg class="glyph stroked chevron-down"><use xlink:href="#stroked-chevron-down"></use></svg></span> Master Lists
				</a>
				<ul class="children collapse" id="sub-item-1">
					<li>
						<a class="" href="account.php?id=<?php echo $id; ?>&masterlists=focus&select=view">
							<svg class="glyph stroked chevron-right"><use xlink:href="#stroked-chevron-right"></use></svg> View
						</a>
					</li>
					<li>
						<a class="" href="account.php?id=<?php echo $id; ?>&masterlists=focus&select=add">
							<svg class="glyph stroked chevron-right"><use xlink:href="#stroked-chevron-right"></use></svg> Import
						</a>
					</li>
					<li>
						<a class="" href="account.php?id=<?php echo $id; ?>&masterlists=focus&select=barcode">
							<svg class="glyph stroked chevron-right"><use xlink:href="#stroked-chevron-right"></use></svg> Settings
						</a>
					</li>
				</ul>
			</li>
			<li class="parent <?php echo $classactivedispatchcontrol ?>">
				<a data-toggle="collapse" href="#sub-item-2">
					<span data-toggle="collapse" href="#sub-item-2"><svg class="glyph stroked chevron-down"><use xlink:href="#stroked-chevron-down"></use></svg></span> Dispatch Control
				</a>
				<ul class="children collapse" id="sub-item-2">
					<li>
						<a class="" href="account.php?id=<?php echo $id; ?>&dispatchcontrol=focus&select=view">
							<svg class="glyph stroked chevron-right"><use xlink:href="#stroked-chevron-right"></use></svg> View
						</a>
					</li>
					<li>
						<a class="" href="account.php?id=<?php echo $id; ?>&dispatchcontrol=focus&select=add">
							<svg class="glyph stroked chevron-right"><use xlink:href="#stroked-chevron-right"></use></svg> Add
						</a>
					</li>
				</ul>
			</li>
			<li class="parent <?php echo $classactiveencodemasterlists ?>">
				<a data-toggle="collapse" href="#sub-item-3">
					<span data-toggle="collapse" href="#sub-item-2"><svg class="glyph stroked chevron-down"><use xlink:href="#stroked-chevron-down"></use></svg></span> Encode Master Lists
				</a>
				<ul class="children collapse" id="sub-item-3">
					<li>
						<a class="" href="account.php?id=<?php echo $id; ?>&encodemasterlists=focus">
							<svg class="glyph stroked chevron-right"><use xlink:href="#stroked-chevron-right"></use></svg> Manual
						</a>
					</li>
					<li>
						<a class="" href="account.php?id=<?php echo $id; ?>&importmasterlists=focus">
							<svg class="glyph stroked chevron-right"><use xlink:href="#stroked-chevron-right"></use></svg> Import
						</a>
					</li>
				</ul>
			</li>
			<li class="parent <?php echo $classactivereports ?>">
				<a data-toggle="collapse" href="#sub-item-4">
					<span data-toggle="collapse" href="#sub-item-2"><svg class="glyph stroked chevron-down"><use xlink:href="#stroked-chevron-down"></use></svg></span> Report
				</a>
				<ul class="children collapse" id="sub-item-4">
					<li>
						<a class="" href="account.php?id=<?php echo $id; ?>&reports=focus">
							<svg class="glyph stroked chevron-right"><use xlink:href="#stroked-chevron-right"></use></svg> Export
						</a>
					</li>
					<li>
						<a class="" href="account.php?id=<?php echo $id; ?>&reportsettings=focus">
							<svg class="glyph stroked chevron-right"><use xlink:href="#stroked-chevron-right"></use></svg> Settings
						</a>
					</li>
				</ul>
			</li>
			<li class="<?php echo $classactiveuploadbackup; ?>"><a href="uploadbackup.php"><svg class="glyph stroked upload"><use xlink:href="#stroked-upload"/></svg> Upload Backup</a></li>
			<li class="<?php echo $classactivesettings; ?>"><a href="account.php?id=<?php echo $id; ?>&settings=focus"><svg class="glyph stroked gear"><use xlink:href="#stroked-gear"/></svg> Account Settings</a></li>
		</ul>
	</div><!--/.sidebar-->
		
	<div class="col-sm-9 col-sm-offset-3 col-lg-10 col-lg-offset-2 main">			
		<div class="row">
			<ol class="breadcrumb">
				<li><a href="#"><svg class="glyph stroked home"><use xlink:href="#stroked-home"></use></svg></a></li>
				<li class="active"><?php echo $pheading; ?></li>
			</ol>
		</div><!--/.row-->
		
		<div class="row">
			<div class="col-lg-12">
				<h1 class="page-header"><?php echo $pheading; ?></h1>
			</div>
		</div><!--/.row-->

		<?php if (isset($_GET["dashboard"]) && trim($_GET["dashboard"]) == "focus") { ?>
			
		<div class="row">
			<div class="col-xs-12 col-md-6 col-lg-3">
				<div class="panel panel-red panel-widget">
					<div class="row no-padding">
						<div class="col-sm-3 col-lg-5 widget-left">
							<svg class="glyph stroked download"><use xlink:href="#stroked-download"/></svg>
						</div>
						<div class="col-sm-9 col-lg-7 widget-right">
							<div class="large">
							<?php
							if ($count_imported_lists < 1000) {
							    // Anything less than a million
							    echo number_format($count_imported_lists);
							} else if ($count_imported_lists < 1000000) {
							    // Anything less than a million
							    echo number_format($count_imported_lists / 1000, 1) . 'K';
							} else if ($count_imported_lists < 1000000000) {
							    // Anything less than a billion
							    echo number_format($count_imported_lists / 1000000, 0) . 'M';
							} else {
							    // At least a billion
							    echo number_format($count_imported_lists / 1000000000, 0) . 'B';
							}
							?>
							</div>
							<div class="text-muted">Imported Lists</div>
						</div>
					</div>
				</div>
			</div>
			<div class="col-xs-12 col-md-6 col-lg-3">
				<div class="panel panel-blue panel-widget ">
					<div class="row no-padding">
						<div class="col-sm-3 col-lg-5 widget-left">
							<svg class="glyph stroked clipboard with paper"><use xlink:href="#stroked-clipboard-with-paper"/></svg>
						</div>
						<div class="col-sm-9 col-lg-7 widget-right">
							<div class="large">
							<?php
							if ($count_record < 1000) {
							    // Anything less than a million
							    echo number_format($count_record);
							} else if ($count_record < 1000000) {
							    // Anything less than a million
							    echo number_format($count_record / 1000, 1) . 'K';
							} else if ($count_record < 1000000000) {
							    // Anything less than a billion
							    echo number_format($count_record / 1000000, 0) . 'M';
							} else {
							    // At least a billion
							    echo number_format($count_record / 1000000000, 0) . 'B';
							}
							?>
							</div>
							<div class="text-muted">Lists Data</div>
						</div>
					</div>
				</div>
			</div>
			<div class="col-xs-12 col-md-6 col-lg-3">
				<div class="panel panel-teal panel-widget">
					<div class="row no-padding">
						<div class="col-sm-3 col-lg-5 widget-left">
							<svg class="glyph stroked male-user"><use xlink:href="#stroked-male-user"></use></svg>
						</div>
						<div class="col-sm-9 col-lg-7 widget-right">
							<div class="large">
							<?php
							if ($count_dispatch_control_messenger < 1000) {
							    // Anything less than a million
							    echo number_format($count_dispatch_control_messenger);
							} else if ($count_dispatch_control_messenger < 1000000) {
							    // Anything less than a million
							    echo number_format($count_dispatch_control_messenger / 1000, 1) . 'K';
							} else if ($count_dispatch_control_messenger < 1000000000) {
							    // Anything less than a billion
							    echo number_format($count_dispatch_control_messenger / 1000000, 0) . 'M';
							} else {
							    // At least a billion
							    echo number_format($count_dispatch_control_messenger / 1000000000, 0) . 'B';
							}
							?>
							</div>
							<div class="text-muted">Dispatch Control</div>
						</div>
					</div>
				</div>
			</div>
			<div class="col-xs-12 col-md-6 col-lg-3">
				<div class="panel panel-orange panel-widget">
					<div class="row no-padding">
						<div class="col-sm-3 col-lg-5 widget-left">
							<svg class="glyph stroked checkmark"><use xlink:href="#stroked-checkmark"/></svg>
						</div>
						<div class="col-sm-9 col-lg-7 widget-right">
							<div class="large">
							<?php
							if ($count_lists_completed < 1000) {
							    // Anything less than a million
							    echo number_format($count_lists_completed);
							} else if ($count_lists_completed < 1000000) {
							    // Anything less than a million
							    echo number_format($count_lists_completed / 1000, 1) . 'K';
							} else if ($count_lists_completed < 1000000000) {
							    // Anything less than a billion
							    echo number_format($count_lists_completed / 1000000, 0) . 'M';
							} else {
							    // At least a billion
							    echo number_format($count_lists_completed / 1000000000, 0) . 'B';
							}
							?>
							</div>
							<div class="text-muted">Lists Completed</div>
						</div>
					</div>
				</div>
			</div>
		</div><!--/.row-->
		

		<?php } else if (isset($_GET["settings"]) && trim($_GET["settings"]) == "focus") { ?>

		<div class="row">
			<div class="col-lg-12">
				<div class="panel panel-default">
					<div class="panel-heading">General</div>
					<div class="panel-body">
						<div class="col-md-6 col-md-offset-3">
							<form role="form" onsubmit="return false;">
							<span id="updateAccStatus"></span>				
								<div class="form-group">
									<input type="text" class="form-control" id="update-acc-email" placeholder="Email" value="<?php echo $email; ?>">
								</div>
								<div class="form-group">
									<input type="password" class="form-control" id="update-acc-pass" placeholder="Password" value="<?php echo $password; ?>">
								</div>
								<button type="submit" class="btn btn-primary" id="updateAccBtn" style="width: 100%" onclick="updateSettings()">Update</button>	
								</form>	
							</div>
					</div>
				</div>
			</div><!-- /.col-->
		</div><!-- /.row -->

		<?php } else if (isset($_GET["masterlists"]) && isset($_GET["select"]) && !isset($_GET["encodeid"]) && trim($_GET["masterlists"]) == "focus" && trim($_GET["select"]) == "add") { ?>


			<div class="row">
				<div class="col-md-12">
				<div class="panel panel-default">
					<div class="panel-heading"><a type="button" class="btn btn-primary" href="account.php?id=<?php echo $id; ?>&masterlists=focus&select=view">View List</a> <a type="button" class="btn btn-primary" href="account.php?id=<?php echo $id; ?>&masterlists=focus&select=barcode">Add Middle Text</a></div>
					<div class="panel-body">
							<form method="POST" id="form-import" action="parsers/upload-excel.php" enctype="multipart/form-data">
								<input id="uploadFile" name="file" class="form-control upload-input" placeholder="File Path.." disabled="disabled"  />
								<div class="fileUpload btn btn-default">
								    <span>Choose File</span>
								    <input id="uploadHiddenBtn" type="file" name="file" class="upload" />
								</div>
								<div class="form-group">
									<select class="form-control" name="import-barcode" id="import-barcode">
										<option value="">Barcode Middle Text</option>
										<option value="manual">*Get the barcode from the masterlist*</option>
										<?php
								  	  	$sql = "SELECT * FROM ppbms_barcode_middle_text WHERE barcode_middle_text_status='1'";
	  									$query = mysqli_query($db_conn, $sql);
	  									$count = 0;
								  	  	while($row = mysqli_fetch_array($query)) {  
										    $count++; 
										    $recid = $row["id"];
										    $code = $row["barcode_middle_text_code"];

										    echo '
										    <option value="'.$code.'">'.$code.'</option>
										    ';   

										}

								
  								 

							  	  		?>
									</select>
								</div>
								<button type="submit" name="submit-excel" id="uploadBtn" disabled="disabled" class="btn btn-primary btn-block">Upload</button>
							</form>
					</div>
					</form>
				</div>
				</div>
			</div><!--/.row-->	

		<?php } else if (isset($_GET["masterlists"]) && isset($_GET["select"]) && !isset($_GET["encodeid"]) && trim($_GET["masterlists"]) == "focus" && trim($_GET["select"]) == "view") { ?>

		   <div class="row">
			<div class="col-md-12">
				<div class="panel panel-default">
					<div class="panel-heading"><a type="button" class="btn btn-primary" href="account.php?id=<?php echo $id; ?>&masterlists=focus&select=add">Import New</a></div>
					<div class="panel-body over-flow">
                    	<div class="date">
					      	<div class="input-group input-append date" id="datePicker">
					            <input type="text" class="form-control" name="date" placeholder="Search (MM/DD/YYYY)" id="search-encode-lists-date"/>
					            <span class="input-group-addon add-on"><span class="glyphicon glyphicon-calendar"></span></span>
					        </div>
					    </div>
					    <br>
					    <div id="encode-lists-result">
                    		<div id="pagination_controls_encode_lists_1"></div>
		                    <div id="record-search-result-encode-lists"></div>
		                    <div id="pagination_controls_encode_lists_2"></div>
		                    <script type="text/javascript">
		                    	request_page_encode_lists(1);
		                    </script>
						 </div>
						</div>
					</div>
				</div>
			</div>
		</div><!--/.row-->	

		<?php } else if (isset($_GET["masterlists"]) && isset($_GET["select"]) && isset($_GET["encodeid"]) && !isset($_GET["editrecid"]) && !isset($_GET["check"]) && trim($_GET["masterlists"]) == "focus" && trim($_GET["select"]) == "view") { ?>

		  <div class="row">
			<div class="col-md-12">
				<div class="panel panel-default">
					<div class="panel-heading"><button type="button" class="btn btn-default" onclick="goBack();">Back</button> <a type="button" class="btn btn-primary" href="account.php?id=<?php echo $id; ?>&masterlists=focus&select=add">Import New</a> <a type="button" class="btn btn-primary" target="_blank" href="receipt.php?id=<?php echo $log_id; ?>&encodeid=<?php echo $_GET["encodeid"]; ?>">Print</a></div>
					<div class="panel-body over-flow">
						<div class="form-group">
							<select class="form-control" id="record-filter-value">
								<option value="all">Filter By All</option>
								<option value="cyclecode">Filter By Cycle Code</option>
								<option value="barcode">Filter By Barcode</option>
								<option value="deltype">Filter By Delivery Type</option>
								<option value="sender">Filter By Sender (Company)</option>
								<option value="name">Filter By Subscriber's Name</option>	
							</select>
						</div>
						<div class="form-group" id="text-search-div">  
                          <input type="text" name="record-search-text" value="" id="record-search-text" placeholder="Search..." class="form-control" autofocus />  
                    	</div>
                    	<div class="form-group" id="deltype-search-div" style="display:none">
							<select class="form-control" id="record-search-deltype">
								<option value="">Delivery Type</option>
								<option value="SOA">SOA</option>
							</select>
						</div>
						<div id="pagination_controls1"></div>
                    	<div id="record-search-result"></div>
                    	<div id="pagination_controls2"></div>
                    	<script type="text/javascript">
                    		request_page_on_load(1);
                    	</script>
					</div>
				</div>
			</div>
		</div><!--/.row-->

		<?php } else if (isset($_GET["masterlists"]) && isset($_GET["select"]) && isset($_GET["encodeid"]) && isset($_GET["editrecid"]) && !isset($_GET["check"]) && trim($_GET["masterlists"]) == "focus" && trim($_GET["select"]) == "view") { ?>

		 	<div class="row">
			<div class="col-lg-12">
				<div class="panel panel-default">
					<div class="panel-heading"><button type="button" class="btn btn-default" onclick="goBack();">Back</button> Edit Master List Record</div>
					<div class="panel-body">
						<?php
							$recid = $_GET["editrecid"];
							$sql = "SELECT * FROM ppbms_record WHERE id='$recid' LIMIT 1";  
  							$result = mysqli_query($db_conn, $sql);  
							while($row = mysqli_fetch_array($result, MYSQLI_ASSOC)) {  
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

							}

						?>
						<div class="col-md-6 col-md-offset-3">
							<form role="form" onsubmit="return false;">
							<span id="updateMasterListsRecordStatus"></span>
								<div class="form-group">
									<label class="control-label">Sender</label>
									<select class="form-control" id="edit-master-list-sender">
										<option value=""></option>
										<?php
											$sql = "SELECT DISTINCT(record_sender) FROM ppbms_record WHERE record_status_status = '1' ORDER BY id DESC";
				  							$query = mysqli_query($db_conn, $sql);
											while($row = mysqli_fetch_array($query)) {

												$rsender = $row["record_sender"];
												$activate = "";  

												if($sender == $rsender) {
													$activate = "selected";
												}

												echo '<option value="'.$row["record_sender"].'" '.$activate.'>'.$row["record_sender"].'</option>';
											}
	  								

										?>	
									</select>
								</div>
								<div class="form-group">
									<label class="control-label">Delivery Type</label>
									<select class="form-control" id="edit-master-list-deltype">
										<option value=""></option>
										<?php
											$sql = "SELECT DISTINCT(record_deltype) FROM ppbms_record WHERE record_status_status = '1' ORDER BY id DESC";
				  							$query = mysqli_query($db_conn, $sql);
											while($row = mysqli_fetch_array($query)) {  

												$rdeltype = $row["record_deltype"];
												$activate = "";  

												if($deltype == $rdeltype) {
													$activate = "selected";
												}


												echo '<option value="'.$row["record_deltype"].'" '.$activate.'>'.$row["record_deltype"].'</option>';
											}

										?>	
									</select>
								</div>
								<div class="form-group">
									<label class="control-label">PUD</label>
					                <div class="date">
					                  <div class="input-group input-append">
					                    <input type="text" class="form-control" value="<?php echo $pud; ?>" name="date" id="edit-pud-date"/>
					                    <span class="input-group-addon add-on"><span class="glyphicon glyphicon-calendar"></span></span>
					                  </div>
					                </div>
					            </div>
					            <div class="form-group">
									<label class="control-label">Month</label>
									<select class="form-control" id="edit-master-list-month">
										<option value=""></option>
										<option value="JANUARY" <?php if($month == "JANUARY") { echo "selected"; } ?>>JANUARY</option>
										<option value="FEBRUARY" <?php if($month == "FEBRUARY") { echo "selected"; } ?>>FEBRUARY</option>
										<option value="MARCH" <?php if($month == "MARCH") { echo "selected"; } ?>>MARCH</option>
										<option value="APRIL" <?php if($month == "APRIL") { echo "selected"; } ?>>APRIL</option>
										<option value="MAY" <?php if($month == "MAY") { echo "selected"; } ?>>MAY</option>
										<option value="JUNE" <?php if($month == "JUNE") { echo "selected"; } ?>>JUNE</option>
										<option value="JULY" <?php if($month == "JULY") { echo "selected"; } ?>>JULY</option>
										<option value="AUGUST" <?php if($month == "AUGUST") { echo "selected"; } ?>>AUGUST</option>
										<option value="SEPTEMBER" <?php if($month == "SEPTEMBER") { echo "selected"; } ?>>SEPTEMBER</option>
										<option value="October" <?php if($month == "OCTOBER") { echo "selected"; } ?>>OCTOBER</option>
										<option value="NOVEMBER" <?php if($month == "NOVEMBER") { echo "selected"; } ?>>NOVEMBER</option>
										<option value="DECEMBER" <?php if($month == "DECEMBER") { echo "selected"; } ?>>DECEMBER</option>
									</select>
								</div>				
								<div class="form-group">
									<label class="control-label">JOB#</label>
									<input type="text" class="form-control" id="edit-master-list-jobnum" value="<?php echo $jobnum; ?>">
								</div>
								<div class="form-group">
									<label class="control-label">Checklist For PPB</label>
									<input type="text" class="form-control" id="edit-checklist-date" value="<?php echo $checklist; ?>">
								</div>
								<div class="form-group">
									<label class="control-label">File Name</label>
									<input type="text" class="form-control" id="edit-master-list-filename" value="<?php echo $filename; ?>">
								</div>
								<div class="form-group">
									<label class="control-label">seq no.</label>
									<input type="text" class="form-control" id="edit-master-list-seqnum" value="<?php echo $seqnum; ?>">
								</div>
								<div class="form-group">
									<label class="control-label">CYCLECODE</label>
									<input type="text" class="form-control" id="edit-master-list-cyclecode" value="<?php echo $cyclecode; ?>">
								</div>
								<div class="form-group">
									<label class="control-label">Qty</label>
									<input type="text" class="form-control" id="edit-master-list-qty" value="<?php echo $qty; ?>">
								</div>
								<div class="form-group">
									<label class="control-label">ADDRESS</label>
									<input type="text" class="form-control" id="edit-master-list-address" value="<?php echo $address; ?>">
								</div>
								<div class="form-group">
									<label class="control-label">AREA</label>
									<input type="text" class="form-control" id="edit-master-list-area" value="<?php echo $area; ?>">
								</div>
								<div class="form-group">
									<label class="control-label">SUBSCRIBERS NAME</label>
									<input type="text" class="form-control" id="edit-master-list-subs" value="<?php echo $subs; ?>">
								</div>
								<div class="form-group">
									<label class="control-label">BARCODE</label>
									<input type="text" class="form-control" id="edit-master-list-barcode" value="<?php echo $barcode; ?>">
								</div>
								<div class="form-group">
									<label class="control-label">ACCT NUMBER</label>
									<input type="text" class="form-control" id="edit-master-list-acctnum" value="<?php echo $acctnum; ?>">
								</div>
								<div class="form-group">
									<label class="control-label">DATE RECEIEVE</label>
					                <div class="date">
					                  <div class="input-group input-append">
					                    <input type="text" class="form-control" name="date" value="<?php echo $datereceive; ?>" id="edit-date-receive-date"/>
					                    <span class="input-group-addon add-on"><span class="glyphicon glyphicon-calendar"></span></span>
					                  </div>
					                </div>
					            </div>
					            <div class="form-group">
									<label class="control-label">RECEIEVE BY</label>
									<input type="text" class="form-control" id="edit-master-list-receiveby" value="<?php echo $receiveby; ?>">
								</div>
								<div class="form-group">
									<label class="control-label">RELATION</label>
									<input type="text" class="form-control" id="edit-master-list-relation" value="<?php echo $relation; ?>">
								</div>
								<div class="form-group">
									<label class="control-label">MESSENGER</label>
									<input type="text" class="form-control" id="edit-master-list-messenger" value="<?php echo $messenger; ?>">
								</div>
								<div class="form-group">
									<label class="control-label">STATUS</label>
									<input type="text" class="form-control" id="edit-master-list-status" value="<?php echo $status; ?>">
								</div>
								<div class="form-group">
									<label class="control-label">Reason for RTS</label>
									<input type="text" class="form-control" id="edit-master-list-reasonrts" value="<?php echo $reasonrts; ?>">
								</div>
								<div class="form-group">
									<label class="control-label">Remarks</label>
									<input type="text" class="form-control" id="edit-master-list-remarks" value="<?php echo $remarks; ?>">
								</div>
								<div class="form-group">
									<label class="control-label">Date Reported</label>
					                <div class="date">
					                  <div class="input-group input-append">
					                    <input type="text" class="form-control" name="date" value="<?php echo $datereported; ?>" id="edit-date-reported-date"/>
					                    <span class="input-group-addon add-on"><span class="glyphicon glyphicon-calendar"></span></span>
					                  </div>
					                </div>
					            </div>
								<button type="submit" class="btn btn-primary" id="updateMasterListsRecordBtn" style="width: 100%" onclick="updateMasterListsRecord(<?php echo $recid;?>)">Update</button>	
								</form>	
							</div>
					</div>
				</div>
			</div><!-- /.col-->
		</div><!-- /.row -->	

		<?php } else if (isset($_GET["masterlists"]) && isset($_GET["select"]) && isset($_GET["encodeid"]) && isset($_GET["check"]) && trim($_GET["masterlists"]) == "focus" && trim($_GET["select"]) == "view" && trim($_GET["check"]) == "unassigned") { ?>

		  <?php
               $encodeid = $_GET["encodeid"];
          ?>

		  <div class="row">
			<div class="col-md-12">
				<div class="panel panel-default">
					<div class="panel-heading"><button type="button" class="btn btn-default" onclick="goBack();">Back</button> <a type="button" class="btn btn-primary" target="_blank" href="receiptnotassign.php?id=<?php echo $log_id; ?>&encodeid=<?php echo $_GET["encodeid"]; ?>">Print</a></div>
					<div class="panel-body over-flow">
						<h2>Item not assign for encode id <?php echo $encodeid; ?></h2>
						<hr>
                    	<div class="form-group" id="deltype-search-div" style="display:none">
							<select class="form-control" id="record-search-deltype">
								<option value="">Delivery Type</option>
								<option value="SOA">SOA</option>
							</select>
						</div>
						<div id="pagination_controls_not_assign_1"></div>
                    	<div id="record-search-result-not-assign"></div>
                    	<div id="pagination_controls_not_assign_2"></div>
                    	<script type="text/javascript">
                    		request_page_not_assign(1, <?php echo $encodeid; ?>);
                    	</script>
					</div>
				</div>
			</div>
		</div><!--/.row-->	

		<?php } else if (isset($_GET["masterlists"]) && isset($_GET["select"]) && !isset($_GET["encodeid"]) && trim($_GET["masterlists"]) == "focus" && trim($_GET["select"]) == "barcode") { ?>

		<div class="row">
			<div class="col-lg-12">
				<div class="panel panel-default">
					<div class="panel-heading"><a type="button" class="btn btn-primary" href="account.php?id=<?php echo $id; ?>&masterlists=focus&select=add">Import New</a> <a type="button" class="btn btn-primary" href="account.php?id=<?php echo $id; ?>&masterlists=focus&select=view">View List</a></div>
					<div class="panel-body">
						<div class="col-md-6">
							<form role="form" onsubmit="return false;">
							<span id="addBarcodeMiddleTextStatus"></span>
								<div class="form-group">
									<label class="control-label">Barcode Middle Text</label>
									<input type="text" class="form-control" id="add-barcode-middle-text">
								</div>
								<button type="submit" class="btn btn-primary" id="addBarcodeMiddleTextBtn" style="width: 100%" onclick="addBarcodeMiddleText()">Add</button>	
								</form>	
						</div>
						<div class="col-md-6">
							<table class="table table-striped">
							  <thead>
								  <tr>
									  <th>No.</th>
									  <th>Code</th>
									  <th>Action</th>
								  </tr>
							  </thead>
							  <tbody>
							  	  	<?php
							  	  	$sql = "SELECT * FROM ppbms_barcode_middle_text WHERE barcode_middle_text_status = '1' ORDER BY id DESC";
  									$query = mysqli_query($db_conn, $sql);
  									$count = 0;
							  	  	while($row = mysqli_fetch_array($query)) {  
									    $count++; 
									    $recid = $row["id"];
									    $code = $row["barcode_middle_text_code"];

									    echo '
									    <tr>
									    <td>'.$count.'</td>		     
								        <td>'.$code.'</td>
								        <td><a href="#" id="'.$recid.'" onclick="openBarcodeMiddleTextUpdateDialog(\''.$recid.'\',\''.$code.'\')">Edit</a> | <a href="#" id="'.$recid.'" onclick="openBarcodeMiddleTextDeleteDialog(\''.$recid.'\')">Delete</a></td>
									    </tr>
									    ';   

									  }

							
  							 

							  	  	?>
						      </tbody>
						  </table>
						</div>
					</div>
				</div>
			</div><!-- /.col-->
		</div><!-- /.row -->

		<?php } else if (isset($_GET["dispatchcontrol"]) && isset($_GET["select"]) && trim($_GET["dispatchcontrol"]) == "focus" && trim($_GET["select"]) == "add") { ?>


			<div class="row">
				<div class="col-md-12">
				<div class="panel panel-default">
					<div class="panel-heading"><a type="button" class="btn btn-primary" href="account.php?id=<?php echo $id; ?>&dispatchcontrol=focus&select=view">View List</a></div>
					<div class="panel-body">
						<div class="col-md-6 col-md-offset-3">
							<form role="form" onsubmit="return false;">
							<span id="addDispatchControlMessengerStatus"></span>
								<div class="form-group">
									<label class="control-label">Messenger Name</label>
									<input type="text" class="form-control" id="add-dispatch-control-messenger-name">
								</div>
								<div class="form-group">
									<label class="control-label">Address</label>
									<input type="text" class="form-control" id="add-dispatch-control-messenger-address">
								</div>
								<div class="form-group">
									<label class="control-label">Prepared By</label>
									<input type="text" class="form-control" id="add-dispatch-control-messenger-prepared">
								</div>
								<div class="form-group">
									<label class="control-label">Date</label>
					                <div class="date">
					                  <div class="input-group input-append">
					                    <input type="text" class="form-control" name="date" id="add-dispatch-control-messenger-date"/>
					                    <span class="input-group-addon add-on"><span class="glyphicon glyphicon-calendar"></span></span>
					                  </div>
					                </div>
					            </div>
								<button type="submit" class="btn btn-primary" id="addDispatchControlMessengerBtn" style="width: 100%" onclick="addDispatchControlMessenger()">Add</button>	
							</form>	
						</div>	
					</div>
				</div>
				</div>
			</div><!--/.row-->	

		<?php } else if (isset($_GET["dispatchcontrol"]) && isset($_GET["select"]) && !isset($_GET["recordidview"]) && !isset($_GET["recordidadd"]) && !isset($_GET["cyclecode"]) && trim($_GET["dispatchcontrol"]) == "focus" && trim($_GET["select"]) == "view") { ?>

		  <div class="row">
			<div class="col-md-12">
				<div class="panel panel-default">
					<div class="panel-heading"><a type="button" class="btn btn-primary" href="account.php?id=<?php echo $id; ?>&dispatchcontrol=focus&select=add">Add New</a></div>
					<div class="panel-body over-flow">
						<div class="form-group" id="text-search-div">  
                          <input type="text" name="dispatch-control-messenger-search-text" value="" id="dispatch-control-messenger-search-text" placeholder="Search..." class="form-control" autofocus /> 
                    	</div>
                    	<div id="dispatch-control-messenger-search-result">
	                    	<div id="pagination_controls_dispatch_1"></div>
	                    	<div id="record-search-result-dispatch"></div>
	                    	<div id="pagination_controls_dispatch_2"></div>
	                    	<script type="text/javascript">
	                    		request_page_dispatch(1);
	                    	</script>
                    	</div>
						</div>
					</div>
				</div>
			</div>
		</div><!--/.row-->	

		<?php } else if (isset($_GET["dispatchcontrol"]) && isset($_GET["select"]) && isset($_GET["recordidview"]) && !isset($_GET["recordidadd"]) && !isset($_GET["cyclecode"]) && trim($_GET["dispatchcontrol"]) == "focus" && trim($_GET["select"]) == "view") { ?>

			<?php
               $recordidview = $_GET["recordidview"];
            ?>

		  <div class="row">
			<div class="col-md-12">
				<div class="panel panel-default">
					<div class="panel-heading"><button type="button" class="btn btn-default" onclick="goBack();">Back</button> <a type="button" class="btn btn-primary" href="account.php?id=<?php echo $log_id; ?>&dispatchcontrol=focus&select=view&recordidadd=<?php echo $recordidview; ?>" id="<?php echo $recordidview; ?>">Add Record to this Messenger</a></div>
					<div class="panel-body over-flow">
						<?php
							$rid = $_GET['recordidview'];
							$sql_messenger = "SELECT * FROM ppbms_dispatch_control_messenger WHERE id = '$rid' ORDER BY id DESC";
  							$query_messenger = mysqli_query($db_conn, $sql_messenger);
							while($row = mysqli_fetch_array($query_messenger)) {  
								$recid = $row["id"];
								$messengername = $row["dispatch_control_messenger_name"];
								$address = $row["dispatch_control_messenger_address"];
								$prepared = $row["dispatch_control_messenger_prepared"];
								$date = $row["dispatch_control_messenger_date"];
								$newDate = date("F d, Y", strtotime($date));
								$newDateUpdate = date("m/d/Y", strtotime($date));
							}

					
  					 

						?>
						<div class="row">
						<div class="col-md-6">
							<h4>Messenger Name: <?php echo $messengername; ?></h4>
						</div>
						<div class="col-md-6">
							<h4>Address: <?php echo $address; ?></h4>
						</div>
						</div>
						<div class="row">
						<div class="col-md-6">
							<h4>Prepared by: <?php echo $prepared; ?></h4> 
						</div>
						<div class="col-md-6">
							<h4>Date: <?php echo $newDate; ?></h4>
						</div>
						</div>
						<hr>
                    	<table class="table table-striped">
							  <thead>
								  <tr>
									  <th>No.</th>
									  <th>Cycle Code</th>
									  <th>Pick-up Date</th>
									  <th>Product Description</th>
									  <th>Unit</th>
									  <th>Quantity</th>
									  <th>Item Received</th>
									  <th>Item Not Received</th>
									  <th>Action</th>
								  </tr>
							  </thead>
							  <tbody>
							  	  	<?php
							  	  	$sql = "SELECT * FROM ppbms_dispatch_control_data WHERE dispatch_control_messenger_id = '$rid' AND dispatch_control_data_status = '1'";
  									$query = mysqli_query($db_conn, $sql);
  									$count = 0;
							  	  	while($row = mysqli_fetch_array($query)) {  
									    $count++; 
									    $recid = $row["id"];
									    $cyclecode = $row["dispatch_control_data_cycle_code"];
									    $date = $row["dispatch_control_data_pickup_date"];
									    $sender = $row["dispatch_control_data_sender"];
									    $deltype = $row["dispatch_control_data_del_type"];
									    $newDate = date("F d, Y", strtotime($date));
									    $newDateUpdate = date("m/d/y", strtotime($date));
									    $newDateUpdate1 = date("n/j/y", strtotime($date));
									    $newDateUpdate2 = date("m/d/Y", strtotime($date));
									    $newDateUpdate3 = date("n/j/Y", strtotime($date));

									    $sql_check_record_count = "SELECT id FROM ppbms_record WHERE messenger_id = '$recordidview' AND record_sender = '$sender' AND record_deltype = '$deltype' AND record_cycle_code = '$cyclecode' AND (record_pud = '$newDateUpdate' OR record_pud = '$newDateUpdate1' OR record_pud = '$newDateUpdate2' OR record_pud = '$newDateUpdate3') AND record_status_status='1'";
									    $query_check_record_count = mysqli_query($db_conn, $sql_check_record_count);
									    $record_count = mysqli_num_rows($query_check_record_count);

									    $sql_check_record_receive = "SELECT id FROM ppbms_record WHERE messenger_id = '$recordidview' AND record_sender = '$sender' AND record_deltype = '$deltype' AND record_cycle_code = '$cyclecode' AND record_status != '' AND (record_pud = '$newDateUpdate' OR record_pud = '$newDateUpdate1') AND record_status_status='1'";
									    $query_check_record_receive = mysqli_query($db_conn, $sql_check_record_receive);
									    $record_count_received = mysqli_num_rows($query_check_record_receive);

									    $sql_check_record_not_receive = "SELECT id FROM ppbms_record WHERE messenger_id = '$recordidview' AND record_sender = '$sender' AND record_deltype = '$deltype' AND record_cycle_code = '$cyclecode' AND record_status = '' AND (record_pud = '$newDateUpdate' OR record_pud = '$newDateUpdate1') AND record_status_status='1'";
									    $query_check_record_not_receive = mysqli_query($db_conn, $sql_check_record_not_receive);
									    $record_count_not_receive = mysqli_num_rows($query_check_record_not_receive);

									    echo '
									    <tr>
									    <td>'.$count.'</td>		     
								        <td>'.$cyclecode.'</td>
								        <td>'.$newDateUpdate.'</td>
								        <td>'.$sender.' '.$deltype.'</td>
								        <td>Pieces</td>
								        <td>'.$record_count.'</td>
								        <td>'.$record_count_received.'</td>
								        <td><a href="account.php?id='.$log_id.'&dispatchcontrol=focus&select=view&cyclecode='.$cyclecode.'&messengerid='.$recordidview.'" id="'.$recid.'">'.$record_count_not_receive.'</a></td>
								        <td><a href="account.php?id='.$log_id.'&dispatchcontrol=focus&select=view&cyclecode='.$cyclecode.'&messengerid='.$recordidview.'&method=manually" id="'.$recid.'">Manually Assign Receipt</a> | <a href="account.php?id='.$log_id.'&dispatchcontrol=focus&select=view&cyclecode='.$cyclecode.'&messengerid='.$recordidview.'&method=import" id="'.$recid.'">Import to Assign Receipt</a></td>
									    </tr>
									    ';   

									  }

							
  							

							  	  	?>
						      </tbody>
						  </table>
					</div>
				</div>
			</div>
		</div><!--/.row-->	

		<?php } else if (isset($_GET["dispatchcontrol"]) && isset($_GET["select"]) && !isset($_GET["recordidview"]) && !isset($_GET["recordidadd"]) && isset($_GET["cyclecode"]) && isset($_GET["messengerid"]) && !isset($_GET["method"]) && trim($_GET["dispatchcontrol"]) == "focus" && trim($_GET["select"]) == "view") { ?>

			<?php
               $cyclecode = $_GET["cyclecode"];
               $messengerid = $_GET["messengerid"];
            ?>

		  <div class="row">
			<div class="col-md-12">
				<div class="panel panel-default">
					<div class="panel-heading"><button type="button" class="btn btn-default" onclick="goBack();">Back</button> <a type="button" class="btn btn-primary" target="_blank" href="receiptnotreceived.php?id=<?php echo $log_id; ?>&messengerid=<?php echo $messengerid; ?>&cyclecode=<?php echo $cyclecode; ?>">Print</a></div>
					<div class="panel-body over-flow">
						<?php
							$sql_messenger = "SELECT * FROM ppbms_dispatch_control_messenger WHERE id = '$messengerid' ORDER BY id DESC";
  							$query_messenger = mysqli_query($db_conn, $sql_messenger);
							while($row = mysqli_fetch_array($query_messenger)) {  
								$recid = $row["id"];
								$messengername = $row["dispatch_control_messenger_name"];
							}

					
  					

						?>
						<h2>Item not received for <?php echo $messengername; ?> with cycle code <?php echo $cyclecode; ?> </h2>
						<hr>
						<div id="pagination_controls_not_recieve_1"></div>
                    	<div id="record-search-result-not-recieve"></div>
                    	<div id="pagination_controls_not_recieve_2"></div>
                    	<script type="text/javascript">
                    		request_page_not_recieve(1, <?php echo $cyclecode; ?>, <?php echo $messengerid; ?>);
                    	</script>
					</div>
				</div>
			</div>
		</div><!--/.row-->	

		<?php } else if (isset($_GET["dispatchcontrol"]) && isset($_GET["select"]) && !isset($_GET["recordidview"]) && !isset($_GET["recordidadd"]) && isset($_GET["cyclecode"]) && isset($_GET["messengerid"]) && isset($_GET["method"]) && !isset($_GET["status"]) && trim($_GET["dispatchcontrol"]) == "focus" && trim($_GET["select"]) == "view" && trim($_GET["method"]) == "manually") { ?>

			<?php
               $cyclecode = $_GET["cyclecode"];
               $messengerid = $_GET["messengerid"];
            ?>

		  <div class="row">
			<div class="col-md-12">
				<div class="panel panel-default">
					<div class="panel-heading"><button type="button" class="btn btn-default" onclick="goBack();">Back</button></div>
					<div class="panel-body over-flow">
						<?php
							$mid = $_GET['messengerid'];
							$sql_messenger = "SELECT * FROM ppbms_dispatch_control_messenger WHERE id = '$mid' ORDER BY id DESC";
  							$query_messenger = mysqli_query($db_conn, $sql_messenger);
							while($row = mysqli_fetch_array($query_messenger)) {  
								$recid = $row["id"];
								$messengername = $row["dispatch_control_messenger_name"];
							}

					
  					

						?>
						<h3>Manually assign receipt/s to <?php echo $messengername; ?> with the cyclecode <?php echo $cyclecode; ?></h3>
						<hr>
						<div class="col-md-6 col-md-offset-3">
						<div class="form-group">
							<label class="control-label">Barcode (You can use barcode scanner or type manually)</label>
							<input type="text" class="form-control" id="barcode-to-assign" autofocus="">
						</div>
						<span id="barcode-to-assign-status"></span>
						</div>
					</div>
				</div>
			</div>
		</div><!--/.row-->	

		<?php } else if (isset($_GET["dispatchcontrol"]) && isset($_GET["select"]) && !isset($_GET["recordidview"]) && !isset($_GET["recordidadd"]) && isset($_GET["cyclecode"]) && isset($_GET["messengerid"]) && isset($_GET["method"]) && !isset($_GET["status"]) && trim($_GET["dispatchcontrol"]) == "focus" && trim($_GET["select"]) == "view" && trim($_GET["method"]) == "import") { ?>

			<?php
               $cyclecode = $_GET["cyclecode"];
               $messengerid = $_GET["messengerid"];
            ?>

            <div class="row">
				<div class="col-md-12">
				<div class="panel panel-default">
					<div class="panel-heading"><button type="button" class="btn btn-default" onclick="goBack();">Back</button></div>
					<div class="panel-body">
						<?php
							$mid = $_GET['messengerid'];
							$sql_messenger = "SELECT * FROM ppbms_dispatch_control_messenger WHERE id = '$mid' ORDER BY id DESC";
  							$query_messenger = mysqli_query($db_conn, $sql_messenger);
							while($row = mysqli_fetch_array($query_messenger)) {  
								$recid = $row["id"];
								$messengername = $row["dispatch_control_messenger_name"];
							}

					
  					

						?>
						<h3>Import to assign receipt/s to <?php echo $messengername; ?> with the cyclecode <?php echo $cyclecode; ?></h3>
						<hr>
							<form method="POST" id="form-import" action="parsers/import-assign-messenger.php" enctype="multipart/form-data">
								<input id="importAssignFile" name="file" class="form-control upload-input" placeholder="File Path.." disabled="disabled"  />
								<div class="fileUpload btn btn-default">
								    <span>Choose File</span>
								    <input id="importAssignHiddenBtn" type="file" name="file" class="upload" />
								</div>
								<div class="form-group">
									<input id="import-assign-sheet-number" type="number" name="import-assign-sheet-number" class="form-control" placeholder="Sheet Number" />
								</div>
								<div class="form-group">
									<input type="hidden" name="import-assign-cycle-code" class="form-control" value="<?php echo $cyclecode; ?>" />
								</div>
								<div class="form-group">
									<input type="hidden" name="import-assign-messenger-id" class="form-control" value="<?php echo $messengerid; ?>" />
								</div>
								<button type="submit" name="import-assign-submit-excel" id="importAssignBtn" disabled="disabled" class="btn btn-primary btn-block">Upload</button>
							</form>
					</div>
				</div>
				</div>
			</div><!--/.row-->	


		<?php } else if (isset($_GET["dispatchcontrol"]) && isset($_GET["select"]) && isset($_GET["importedcount"]) && !isset($_GET["recordidview"]) && !isset($_GET["recordidadd"]) && isset($_GET["cyclecode"]) && isset($_GET["messengerid"]) && isset($_GET["method"]) && isset($_GET["status"]) && trim($_GET["dispatchcontrol"]) == "focus" && trim($_GET["select"]) == "view" && trim($_GET["method"]) == "import" && trim($_GET["status"]) == "success") { ?>

			<?php
               $cyclecode = $_GET["cyclecode"];
               $messengerid = $_GET["messengerid"];
            ?>

            <div class="row">
				<div class="col-md-12">
				<div class="panel panel-default">
					<div class="panel-heading"><a type="button" class="btn btn-default" href="account.php?id=<?php echo $id; ?>&dispatchcontrol=focus&select=view&recordidview=<?php echo $messengerid; ?>">Back</a></div>
					<div class="panel-body">
						<h3 style="color: #00ee00;"><?php echo $_GET["importedcount"]; ?> row/s successfully imported and assigned!</h3>
					</div>
				</div>
				</div>
			</div><!--/.row-->	
		  

		<?php } else if (isset($_GET["dispatchcontrol"]) && isset($_GET["select"]) && !isset($_GET["recordidview"]) && !isset($_GET["cyclecode"]) && isset($_GET["recordidadd"]) && trim($_GET["dispatchcontrol"]) == "focus" && trim($_GET["select"]) == "view") { ?>

		  <div class="row">
			<div class="col-md-12">
				<div class="panel panel-default">
					<div class="panel-heading"><button type="button" class="btn btn-default" onclick="goBack();">Back</button></div>
					<div class="panel-body over-flow">
					<div class="col-md-6 col-md-offset-3">
						<form id="loginForm" onsubmit="return false;">
				<span id="addDispatchControlDataStatus"></span>
				<input type="hidden" class="form-control" id="add-dispatch-control-data-messenger-id">
									
								<div class="form-group">
									<label class="control-label">Cycle Code</label>
										<select class="form-control" id="add-dispatch-control-data-cycle-code">
											<option value=""></option>
											<?php
												$sql = "SELECT DISTINCT(record_cycle_code) FROM ppbms_record WHERE record_status_status = '1' ORDER BY id DESC LIMIT 60";
					  							$query = mysqli_query($db_conn, $sql);
												while($row = mysqli_fetch_array($query)) {  
													echo '<option value="'.$row["record_cycle_code"].'">'.$row["record_cycle_code"].'</option>';
												}

										
  										

											?>	
										</select>
								</div>
								<div class="form-group">
									<label class="control-label">Pick-up Date</label>
					                <div class="date">
					                  <div class="input-group input-append">
					                    <input type="text" class="form-control" name="date" id="add-dispatch-control-data-pickup-date"/>
					                    <span class="input-group-addon add-on"><span class="glyphicon glyphicon-calendar"></span></span>
					                  </div>
					                </div>
					            </div>
								<div class="form-group">
									<label class="control-label">Sender</label>
										<select class="form-control" id="add-dispatch-control-data-sender">
											<option value=""></option>
											<?php
												$sql = "SELECT DISTINCT(record_sender) FROM ppbms_record WHERE record_status_status = '1' ORDER BY id DESC";
					  							$query = mysqli_query($db_conn, $sql);
												while($row = mysqli_fetch_array($query)) {  
													echo '<option value="'.$row["record_sender"].'">'.$row["record_sender"].'</option>';
												}

										
  										

											?>	
										</select>
								</div>
								<div class="form-group">
									<label class="control-label">Delivery Type</label>
										<select class="form-control" id="add-dispatch-control-data-del-type">
											<option value=""></option>
											<?php
												$sql = "SELECT DISTINCT(record_deltype) FROM ppbms_record WHERE record_status_status = '1' ORDER BY id DESC";
					  							$query = mysqli_query($db_conn, $sql);
												while($row = mysqli_fetch_array($query)) {  
													echo '<option value="'.$row["record_deltype"].'">'.$row["record_deltype"].'</option>';
												}

  										

											?>	
										</select>
								</div>
						<button type="button" class="btn btn-primary" id="addDispatchControlDataBtn" style="width: 100%" onclick="addDispatchControlDataRecord()">Add</button>
			</form>
			</div>
					</div>
				</div>
			</div>
		</div><!--/.row-->	

		<?php } else if (isset($_GET["encodemasterlists"]) && trim($_GET["encodemasterlists"]) == "focus") { ?>

			<div class="row">
				<div class="col-md-12">
				<div class="panel panel-default">
					<div class="panel-heading"><a type="button" class="btn btn-primary" href="account.php?id=<?php echo $id; ?>&masterlists=focus&select=view">View Master Lists</a></div>
					<div class="panel-body">
						<div class="col-md-6 col-md-offset-3">
							<form id="loginForm" onsubmit="return false;">
								<div class="form-group">
									<label class="control-label">Barcode (You can use barcode scanner or type manually)</label>
									<input type="text" class="form-control" id="barcode-to-encode-search" autofocus="">
								</div>
								<button type="button" class="btn btn-primary" id="addDispatchControlDataBtn" style="width: 100%" onclick="resetToEncodeResult()">Reset</button>
							</form>
						</div>
						<div class="col-md-12">
						<hr>
						<div id="to-encode-result" style="display: none;">
							<form id="loginForm" onsubmit="return false;">
    <span id="encodeMasterListsStatus"></span>
    <input type="hidden" class="form-control" id="encode-master-lists-id">
    <div class="form-group">
    <label class="control-label">Date Recieved</label>
    <div class="date">
    <div class="input-group input-append">
    <input type="text" class="form-control" name="date" id="update-record-date-recieve"/>
    <span class="input-group-addon add-on"><span class="glyphicon glyphicon-calendar"></span></span>
    </div>
    </div>
    </div>
    <div class="form-group">
    <label class="control-label">Recieve By</label>
    <input type="text" class="form-control" id="update-record-receive-by">
    </div>
    <div class="form-group">
    <label class="control-label">Relation</label>
    <input type="text" class="form-control" id="update-record-relation">
    </div>
    <div class="form-group">
    <label class="control-label">Messenger</label>
    <input type="text" class="form-control" id="update-record-messenger">
    </div>
    <div class="form-group">
    <label class="control-label">Status</label>
    <input type="text" class="form-control" id="update-record-status">
    </div>
    <div class="form-group">
    <label class="control-label">Reason RTS</label>
    <input type="text" class="form-control" id="update-record-reason-rts">
    </div>
    <div class="form-group">
    <label class="control-label">Remarks</label>
    <input type="text" class="form-control" id="update-record-remarks">
    </div>
    <div class="form-group">
    <label class="control-label">Date Reported</label>
    <div class="date">
    <div class="input-group input-append">
    <input type="text" class="form-control" name="date" id="update-record-date-reported"/>
    <span class="input-group-addon add-on"><span class="glyphicon glyphicon-calendar"></span></span>
    </div>
    </div>
    </div>
    <button type="button" class="btn btn-primary" id="encodeMasterListsBtn" style="width: 100%" onclick="encodeMasterListsRecord()">Add</button>
    </form>
						</div>
						<div id="to-encode-result-no-data-found" style="display: none;">
							<center><h4>No data found!</h4></center>
						</div>
						</div>
					</div>
					</form>
				</div>
				</div>
			</div><!--/.row-->	

		<?php } else if (isset($_GET["importmasterlists"]) && !isset($_GET["status"]) && trim($_GET["importmasterlists"]) == "focus") { ?>

			<div class="row">
				<div class="col-md-12">
				<div class="panel panel-default">
					<div class="panel-heading"><a type="button" class="btn btn-primary" href="account.php?id=<?php echo $id; ?>&masterlists=focus&select=view">View Master Lists</a></div>
					<div class="panel-body">
							<form method="POST" id="form-import" action="parsers/import-encode-upload-excel.php" enctype="multipart/form-data">
								<input id="importEncodeFile" name="file" class="form-control upload-input" placeholder="File Path.." disabled="disabled"  />
								<div class="fileUpload btn btn-default">
								    <span>Choose File</span>
								    <input id="importEncodeHiddenBtn" type="file" name="file" class="upload" />
								</div>
								<div class="form-group">
									<input id="import-encode-sheet-number" type="number" name="import-encode-sheet-number" class="form-control" placeholder="Sheet Number" />
								</div>
								<button type="submit" name="import-encode-submit-excel" id="importEncodeBtn" disabled="disabled" class="btn btn-primary btn-block">Upload</button>
							</form>
					</div>
					</form>
				</div>
				</div>
			</div><!--/.row-->

		<?php } else if (isset($_GET["importmasterlists"]) && isset($_GET["status"]) && isset($_GET["importedcount"]) && trim($_GET["importmasterlists"]) == "focus" && trim($_GET["status"]) == "success") { ?>

			<div class="row">
				<div class="col-md-12">
				<div class="panel panel-default">
					<div class="panel-heading"><a type="button" class="btn btn-default" href="account.php?id=<?php echo $id; ?>&importmasterlists=focus">Back</a></div>
					<div class="panel-body">
						<h3 style="color: #00ee00;"><?php echo $_GET["importedcount"]; ?> row/s has successfully imported and encoded!</h3>
					</div>
				</div>
				</div>
			</div><!--/.row-->	


		<?php } else if (isset($_GET["reports"]) && trim($_GET["reports"]) == "focus") { ?>

			 <div class="row">
				<div class="col-md-12">
					<div class="panel panel-default">
						<div class="panel-heading"><a type="button" class="btn btn-primary" href="account.php?id=<?php echo $id; ?>&reportsettings=focus">Report Settings</a></div>
						<div class="panel-body over-flow">
	                    	<div class="form-group">
								<select class="form-control" id="report-sender">
									<option value="">Sender</option>
									<?php
										$sql = "SELECT DISTINCT(record_sender) FROM ppbms_record WHERE record_status_status = '1' ORDER BY id DESC";
			  							$query = mysqli_query($db_conn, $sql);
										while($row = mysqli_fetch_array($query)) {  
											echo '<option value="'.$row["record_sender"].'">'.$row["record_sender"].'</option>';
										}
  								

									?>	
								</select>
							</div>
							<div class="form-group">
								<select class="form-control" id="report-deltype">
									<option value="">Delivery Type</option>
									<?php
										$sql = "SELECT DISTINCT(record_deltype) FROM ppbms_record WHERE record_status_status = '1' ORDER BY id DESC";
			  							$query = mysqli_query($db_conn, $sql);
										while($row = mysqli_fetch_array($query)) {  
											echo '<option value="'.$row["record_deltype"].'">'.$row["record_deltype"].'</option>';
										}

									?>	
								</select>
							</div>
							<div class="form-group">
								<select class="form-control" id="report-month">
									<option value="">Month</option>
									<option value="January">January</option>
									<option value="February">February</option>
									<option value="March">March</option>
									<option value="April">April</option>
									<option value="May">May</option>
									<option value="June">June</option>
									<option value="July">July</option>
									<option value="August">August</option>
									<option value="September">September</option>
									<option value="October">October</option>
									<option value="November">November</option>
									<option value="December">December</option>
								</select>
							</div>
							<div class="form-group">
								<select class="form-control" id="report-year">
									<option value="">Year</option>
									<option value="2017">2017</option>
									<option value="2018">2018</option>
									<option value="2019">2019</option>
									<option value="2020">2020</option>
									<option value="2021">2021</option>
									<option value="2022">2022</option>
									<option value="2023">2023</option>
									<option value="2024">2024</option>
									<option value="2025">2025</option>
									<option value="2026">2026</option>
									<option value="2027">2027</option>
									<option value="2028">2028</option>
									<option value="2021">2029</option>
									<option value="2022">2030</option>
									<option value="2023">2031</option>
									<option value="2024">2032</option>
									<option value="2025">2033</option>
									<option value="2026">2034</option>
									<option value="2027">2035</option>
									<option value="2028">2036</option>
									<option value="2027">2037</option>
									<option value="2028">2038</option>
								</select>
							</div>
							<hr>
							<div id="report-checkbox-list"></div>
						</div>
						</div>
					</div>
				</div>
			</div><!--/.row-->	

		<?php } else if (isset($_GET["reportsettings"]) && trim($_GET["reportsettings"]) == "focus") { ?>

		<div class="row">
			<div class="col-lg-12">
				<div class="panel panel-default">
					<div class="panel-heading"><a type="button" class="btn btn-primary" href="account.php?id=<?php echo $id; ?>&reports=focus">Export</a></div>
					<div class="panel-body">
						<div class="col-md-12">
							<table class="table table-striped">
							  <thead>
								  <tr>
									  <th>No.</th>
									  <th>Area</th>
									  <th>Price</th>
									  <th>Action</th>
								  </tr>
							  </thead>
							  <tbody>
							  	  	<?php
							  	  	$sql = "SELECT * FROM ppbms_area_price WHERE area_price_status = '1'";
  									$query = mysqli_query($db_conn, $sql);
  									$count = 0;
							  	  	while($row = mysqli_fetch_array($query)) {  
									    $count++; 
									    $recid = $row["id"];
									    $name = $row["area_price_name"];
									    $price = $row["area_price_price"];

									    echo '
									    <tr>
									    <td>'.$count.'</td>		     
								        <td>'.$name.'</td>
								        <td>'.$price.'</td>
								        <td><a href="#" id="'.$recid.'" onclick="openAreaPriceUpdateDialog(\''.$recid.'\',\''.$name.'\',\''.$price.'\')">Edit</a></td>
									    </tr>
									    ';   

									  }

									
  									

							  	  	?>
						      </tbody>
						  </table>
						</div>
					</div>
				</div>
			</div><!-- /.col-->
		</div><!-- /.row -->

		<?php } ?>

	</div>	
	<!--/.main-->

	<!-- Modal -->
	<div class="modal fade bd-example-modal-sm" id="editBarcodeMiddleText" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
	  <div class="modal-dialog modal-sm" role="document">
	    <div class="modal-content">
	      <div class="modal-header">
	        <h5 class="modal-title" id="exampleModalLabel">Edit Barcode Middle Text</h5>
	      </div>
	      <div class="modal-body">
			<form id="loginForm" onsubmit="return false;">
				<span id="updateBarcodeMiddleTextStatus"></span>
				<input type="hidden" class="form-control" id="update-barcode-middle-text-id">
				<div class="row">
				  <div class="col-md-12">
	              <div class="form-group">
	                <label for="username" class="control-label">Barcode Middle Text</label>
	                <input type="text" class="form-control" id="update-barcode-middle-text-code" name="code">
	                <span class="help-block"></span>
	              </div>
				</div>
				</div>
			</form>
	      </div>
	      <div class="modal-footer">
	        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
	        <button type="button" class="btn btn-primary" id="updateBarcodeMiddleTextBtn" onclick="updateBarcodeMiddleTextRecord()">Update</button>
	      </div>
	    </div>
	  </div>
	</div>

	<!-- Modal -->
	<div class="modal fade bd-example-modal-sm" id="editAreaPrice" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
	  <div class="modal-dialog modal-sm" role="document">
	    <div class="modal-content">
	      <div class="modal-header">
	        <h5 class="modal-title" id="exampleModalLabel">Edit Area Price</h5>
	      </div>
	      <div class="modal-body">
			<form id="loginForm" onsubmit="return false;">
				<span id="updateAreaPriceStatus"></span>
				<input type="hidden" class="form-control" id="update-area-price-id">
				<div class="row">
				  <div class="col-md-12">
	              <div class="form-group">
	                <label for="username" class="control-label">Name</label>
	                <input type="text" class="form-control" id="update-area-price-name" name="code" disabled>
	                <span class="help-block"></span>
	              </div>
	              <div class="form-group">
	                <label for="username" class="control-label">Price</label>
	                <input type="text" class="form-control" id="update-area-price-price" name="code">
	                <span class="help-block"></span>
	              </div>
				</div>
				</div>
			</form>
	      </div>
	      <div class="modal-footer">
	        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
	        <button type="button" class="btn btn-primary" id="updateAreaPriceBtn" onclick="updateAreaPriceRecord()">Update</button>
	      </div>
	    </div>
	  </div>
	</div>

	<!-- Modal -->
	<div class="modal fade bd-example-modal-sm" id="deleteBarcodeMiddleText" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
	  <div class="modal-dialog modal-sm" role="document">
	    <div class="modal-content">
	      <div class="modal-header">
	        <h5 class="modal-title" id="exampleModalLabel">Delete Barcode Middle Text</h5>
	      </div>
	      <div class="modal-body">
	      	<span id="deleteBarcodeMiddleTextStatus"></span>
	        <input type="hidden" class="form-control" id="delete-barcode-middle-text-id">
	        <p>Are you sure with this?</p>
	      </div>
	      <div class="modal-footer">
	        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
	        <button type="button" class="btn btn-primary" id="deleteBarcodeMiddleTextBtn" onclick="deleteBarcodeMiddleTextRecord()">Yes</button>
	      </div>
	    </div>
	  </div>
	</div> 

	<!-- Modal -->
	<div class="modal fade bd-example-modal-sm" id="deleteMasterListItem" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
	  <div class="modal-dialog modal-sm" role="document">
	    <div class="modal-content">
	      <div class="modal-header">
	        <h5 class="modal-title" id="exampleModalLabel">Delete Master List Item</h5>
	      </div>
	      <div class="modal-body">
	      	<span id="deleteMasterListItemStatus"></span>
	        <input type="hidden" class="form-control" id="delete-master-list-item-id">
	        <p>Are you sure with this?</p>
	      </div>
	      <div class="modal-footer">
	        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
	        <button type="button" class="btn btn-primary" id="deleteMasterListItemBtn" onclick="deleteMasterListItemRecord()">Yes</button>
	      </div>
	    </div>
	  </div>
	</div> 

	<!-- Modal -->
	<div class="modal fade bd-example-modal-sm" id="editDispatchControlMessenger" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
	  <div class="modal-dialog modal-sm" role="document">
	    <div class="modal-content">
	      <div class="modal-header">
	        <h5 class="modal-title" id="exampleModalLabel">Edit Dispatch Control Messenger</h5>
	      </div>
	      <div class="modal-body">
			<form id="loginForm" onsubmit="return false;">
				<span id="updateDispatchControlMessengerStatus"></span>
				<input type="hidden" class="form-control" id="update-dispatch-control-messenger-id">
				<div class="row">
				  <div class="col-md-12">
								<div class="form-group">
									<label class="control-label">Messenger Name</label>
									<input type="text" class="form-control" id="update-dispatch-control-messenger-name">
								</div>
								<div class="form-group">
									<label class="control-label">Address</label>
									<input type="text" class="form-control" id="update-dispatch-control-messenger-address">
								</div>
								<div class="form-group">
									<label class="control-label">Prepared By</label>
									<input type="text" class="form-control" id="update-dispatch-control-messenger-prepared">
								</div>
								<div class="form-group">
									<label class="control-label">Date</label>
					                <div class="date">
					                  <div class="input-group input-append date" id="datePicker">
					                    <input type="text" class="form-control" name="date" id="update-dispatch-control-messenger-date"/>
					                    <span class="input-group-addon add-on"><span class="glyphicon glyphicon-calendar"></span></span>
					                  </div>
					                </div>
					            </div>
				</div>
				</div>
			</form>
	      </div>
	      <div class="modal-footer">
	        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
	        <button type="button" class="btn btn-primary" id="updateDispatchControlMessengerBtn" onclick="updateDispatchControlMessengerRecord()">Update</button>
	      </div>
	    </div>
	  </div>
	</div>

	<!-- Modal -->
	<div class="modal fade bd-example-modal-sm" id="deleteDispatchControlMessenger" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
	  <div class="modal-dialog modal-sm" role="document">
	    <div class="modal-content">
	      <div class="modal-header">
	        <h5 class="modal-title" id="exampleModalLabel">Delete Barcode Middle Text</h5>
	      </div>
	      <div class="modal-body">
	      	<span id="deleteDispatchControlMessengerStatus"></span>
	        <input type="hidden" class="form-control" id="delete-dispatch-control-messenger-id">
	        <p>Are you sure with this?</p>
	      </div>
	      <div class="modal-footer">
	        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
	        <button type="button" class="btn btn-primary" id="deleteDispatchControlMessengerBtn" onclick="deleteDispatchControlMessengerRecord()">Yes</button>
	      </div>
	    </div>
	  </div>
	</div> 

	<!-- Modal -->
	<div class="modal fade bd-example-modal-sm" id="deleteEncodeLists" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
	  <div class="modal-dialog modal-sm" role="document">
	    <div class="modal-content">
	      <div class="modal-header">
	        <h5 class="modal-title" id="exampleModalLabel">Delete Master Lists</h5>
	      </div>
	      <div class="modal-body">
	      	<span id="deleteEncodeListsStatus"></span>
	        <input type="hidden" class="form-control" id="delete-encode-lists-id">
	        <p>Are you sure with this?</p>
	      </div>
	      <div class="modal-footer">
	        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
	        <button type="button" class="btn btn-primary" id="deleteEncodeListsBtn" onclick="deleteEncodeListsRecord()">Yes</button>
	      </div>
	    </div>
	  </div>
	</div> 

		<!-- Modal -->
	<div class="modal fade bd-example-modal-lg" id="viewDispatchControlData" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
	  <div class="modal-dialog modal-lg" role="document">
	    <div class="modal-content">
	      <div class="modal-header">
	        <h5 class="modal-title" id="exampleModalLabel">Dispatch Control Data</h5>
	      </div>
	      <div class="modal-body">

	      </div>
	      <div class="modal-footer">
	        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
	      </div>
	    </div>
	  </div>
	</div>

	<script src="js/jquery-1.11.1.min.js"></script>
	<script src="js/bootstrap.min.js"></script>
<!-- 	<script src="js/chart.min.js"></script>
	<script src="js/chart-data.js"></script>
	<script src="js/easypiechart.js"></script>
	<script src="js/easypiechart-data.js"></script>-->
<!-- 	<script src="js/bootstrap-datepicker.js"></script>  -->
	<script src="js/bootstrap-table.js"></script>
	<script src="js/main.js"></script>
    <script src="js/ajax.js"></script>
    <script src="//cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.3.0/js/bootstrap-datepicker.min.js"></script>

    <?php if(isset($_GET["select"]) && $_GET["select"] == "add") { ?>
	  <script type="text/javascript">
	   	document.getElementById("uploadHiddenBtn").onchange = function () {
		    document.getElementById("uploadFile").value = this.value;
		};
	  </script>
  	<?php } else if (isset($_GET["importmasterlists"]) && $_GET["importmasterlists"] == "focus") { ?>
	  <script type="text/javascript">
	   	document.getElementById("importEncodeHiddenBtn").onchange = function () {
		    document.getElementById("importEncodeFile").value = this.value;
		};
	  </script>
	  <?php } else if (isset($_GET["method"]) && !isset($_GET["status"]) && $_GET["method"] == "import") { ?>
	  <script type="text/javascript">
	   	document.getElementById("importAssignHiddenBtn").onchange = function () {
		    document.getElementById("importAssignFile").value = this.value;
		};
	  </script>
	 <?php } ?>

	<script>

		!function ($) {
		    $(document).on("click","ul.nav li.parent > a > span.icon", function(){          
		        $(this).find('em:first').toggleClass("glyphicon-minus");      
		    }); 
		    $(".sidebar span.icon").find('em:first').addClass("glyphicon-plus");
		}(window.jQuery);

		$(window).on('resize', function () {
		  if ($(window).width() > 768) $('#sidebar-collapse').collapse('show')
		})
		$(window).on('resize', function () {
		  if ($(window).width() <= 767) $('#sidebar-collapse').collapse('hide')
		})

		function updateSettings() {
        var email = _("update-acc-email").value;
        var password = _("update-acc-pass").value;
        var load = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i><center>';
        var error = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
        _("updateAccStatus").innerHTML = load;
          if (email == "" || password == "") {
                _("updateAccStatus").innerHTML = error;
          } else {
                _("updateAccBtn").disabled = true;
                _("updateAccStatus").innerHTML = load;
                var ajax = ajaxObj("POST", "parsers/account.php");
                ajax.onreadystatechange = function() {
                  if(ajaxReturn(ajax) == true) {
                      if (ajax.responseText == "successinsert"){
                        var x = document.getElementById("snackbar-update-success")
                        x.className = "show";
                        setTimeout(function(){ x.className = x.className.replace("show", ""); }, 3000);
                        _("updateAccBtn").disabled = false;
                        _("updateAccStatus").innerHTML = "";
                      } else if (ajax.responseText == "samedata"){
                        _("updateAccBtn").disabled = false;
                        _("updateAccStatus").innerHTML = '<div style="color: red; margin-bottom: 10px;">No changes detected!</div>';
                      } else {
                          _("updateAccBtn").disabled = false;
                          _("updateAccStatus").innerHTML = '<div style="color: red; margin-bottom: 10px;">User does not exist!</div>';
                      }
                  }
                }
          ajax.send("updateemail="+email+"&password="+password);
          }
      }
      function updateMasterListsRecord(recid) {
        var sender = _("edit-master-list-sender").value;
        var deltype = _("edit-master-list-deltype").value;
        var pud = _("edit-pud-date").value;
        var month = _("edit-master-list-month").value;
        var jobnum = _("edit-master-list-jobnum").value;
        var checklist = _("edit-checklist-date").value;
        var filename = _("edit-master-list-filename").value;
        var seqnum = _("edit-master-list-seqnum").value;
        var cyclecode = _("edit-master-list-cyclecode").value;
        var qty = _("edit-master-list-qty").value;
        var address = _("edit-master-list-address").value;
        var area = _("edit-master-list-area").value;
        var subs = _("edit-master-list-subs").value;
        var barcode = _("edit-master-list-barcode").value;
        var acctnum = _("edit-master-list-acctnum").value;
        var receive = _("edit-date-receive-date").value;
        var receiveby = _("edit-master-list-receiveby").value;
        var relation = _("edit-master-list-relation").value;
        var messenger = _("edit-master-list-messenger").value;
        var status = _("edit-master-list-status").value;
        var reasonrts = _("edit-master-list-reasonrts").value;
        var remarks = _("edit-master-list-remarks").value;
        var reported = _("edit-date-reported-date").value;
        var load = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i><center>';
        var error = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
        _("updateMasterListsRecordStatus").innerHTML = load;
          if (sender == "" || deltype == "" || pud == "" || month == "" || jobnum == "" || checklist == "" || filename == "" || seqnum == "" || cyclecode == "" || qty == "" || address == "" || area == "" || subs == "" || barcode == "") {
                _("updateMasterListsRecordStatus").innerHTML = error;
          } else {
                _("updateMasterListsRecordBtn").disabled = true;
                _("updateMasterListsRecordStatus").innerHTML = load;
                var ajax = ajaxObj("POST", "parsers/account.php");
                ajax.onreadystatechange = function() {
                  if(ajaxReturn(ajax) == true) {
                      if (ajax.responseText == "successupdate"){
                        location.reload();
                      } else if (ajax.responseText == "error"){
                        _("updateMasterListsRecordBtn").disabled = false;
                        _("updateMasterListsRecordStatus").innerHTML = '<div style="color: red; margin-bottom: 10px;">Record does not exist!</div>';
                      } else {
                          _("updateMasterListsRecordBtn").disabled = false;
                          _("updateMasterListsRecordStatus").innerHTML = '<div style="color: red; margin-bottom: 10px;">Unknown Error!</div>';
                      }
                  }
                }
          ajax.send("updatemasterlistrid="+recid+"&updatemasterlistsender="+sender+"&updatemasterlistdeltype="+deltype+"&updatemasterlistpud="+pud+"&updatemasterlistmonth="+month+"&updatemasterlistjobnum="+jobnum+"&updatemasterlistchecklist="+encodeURIComponent(checklist)+"&updatemasterlistfilename="+encodeURIComponent(filename)+"&updatemasterlistseqnum="+seqnum+"&updatemasterlistcyclecode="+cyclecode+"&updatemasterlistqty="+qty+"&updatemasterlistaddress="+encodeURIComponent(address)+"&updatemasterlistarea="+area+"&updatemasterlistsubs="+subs+"&updatemasterlistbarcode="+barcode+"&updatemasterlistacctnum="+acctnum+"&updatemasterlistreceive="+encodeURIComponent(receive)+"&updatemasterlistreceiveby="+receiveby+"&updatemasterlistrelation="+relation+"&updatemasterlistmessenger="+messenger+"&updatemasterliststatus="+status+"&updatemasterlistreasonrts="+reasonrts+"&updatemasterlistreasonremarks="+encodeURIComponent(remarks)+"&updatemasterlistreasonreported="+reported);
          }
      }
      function addRecord() {
      	var date = _("add-date").value;
        var elco = _("add-el-co").value;
        var elapp = _("add-el-app").value;
        var sponcom = _("add-spon-com").value;
        var gasera = _("add-gasera").value;
        var staele = _("add-sta-ele").value;
        var lpg = _("add-lpg").value;
        var chem = _("add-chem").value;
        var phyro = _("add-phyro").value;
        var matlig = _("add-mat-lig").value;
        var light = _("add-light").value;
        var exp = _("add-exp").value;
        var resfire = _("add-res-fire").value;
        var edufire = _("add-edu-fire").value;
        var hcfire = _("add-hc-fire").value;
        var stofire = _("add-sto-fire").value;
        var bizfire = _("add-biz-fire").value;
        var mixfire = _("add-mix-fire").value;
        var indfire = _("add-ind-fire").value;
        var commer = _("add-com-mer").value;
        var assembly = _("add-assembly").value;
        var others = _("add-others").value;
        var grass = _("add-grass").value;
        var post = _("add-post").value;
        var veh = _("add-veh").value;
        var fatci = _("add-fat-ci").value;
        var fatbf = _("add-fat-bf").value;
        var injci = _("add-inj-ci").value;
        var injbf = _("add-inj-bf").value;
        var load = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i><center>';
        var error = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
        _("addRecordStatus").innerHTML = load;
          if (date == "" || elco == "" || elco == "0" || elapp == "" || elapp == "0" || sponcom == "" || sponcom == "0" || gasera == "" || gasera == "0" || staele == "" || staele == "0" || lpg == "" || lpg == "0" || chem == "" || chem == "0" || phyro == "" || phyro == "0" || matlig == "" || matlig == "0" || light == "" || light == "0" || exp == "" || exp == "0" || resfire == "" || resfire == "0" || edufire == "" || edufire == "0" || hcfire == "" || hcfire == "0" || stofire == "" || stofire == "0" || bizfire == "" || bizfire == "0" || mixfire == "" || mixfire == "0" || indfire == "" || indfire == "0" || commer == "" || commer == "0" || assembly == "" || assembly == "0" || others == "" || others == "0" || grass == "" || grass == "0" || post == "" || post == "0" || veh == "" || veh == "0" || fatci == "" || fatci == "0" || fatbf == "" || fatbf == "0" || injci == "" || injci == "0" || injbf == "" || injbf == "0") {
                _("addRecordStatus").innerHTML = error;
          } else {
                _("addRecordBtn").disabled = true;
                _("addRecordStatus").innerHTML = load;
                var ajax = ajaxObj("POST", "parsers/account.php");
                ajax.onreadystatechange = function() {
                  if(ajaxReturn(ajax) == true) {
                      if (ajax.responseText == "successinsert"){
                        _("add-date").value = "";
                      	_("add-el-co").value = "";
                      	_("add-el-app").value = "";
                      	_("add-spon-com").value = "";
                      	_("add-gasera").value = "";
                      	_("add-sta-ele").value = "";
                      	_("add-lpg").value = "";
                      	_("add-chem").value = "";
                      	_("add-phyro").value = "";
                      	_("add-mat-lig").value = "";
                      	_("add-light").value = "";
                      	_("add-exp").value = "";
                      	_("add-res-fire").value = "";
                      	_("add-edu-fire").value = "";
                      	_("add-hc-fire").value = "";
                      	_("add-sto-fire").value = "";
                      	_("add-biz-fire").value = "";
                      	_("add-mix-fire").value = "";
                      	_("add-ind-fire").value = "";
                      	_("add-com-mer").value = "";
                      	_("add-assembly").value = "";
                      	_("add-others").value = "";
                      	_("add-grass").value = "";
                      	_("add-post").value = "";
                      	_("add-veh").value = "";
                      	_("add-fat-ci").value = "";
                      	_("add-fat-bf").value = "";
                      	_("add-inj-ci").value = "";
                      	_("add-inj-bf").value = "";
                        var x = document.getElementById("snackbar-add-success")
                        x.className = "show";
                        setTimeout(function(){ x.className = x.className.replace("show", ""); }, 3000);
                        _("addRecordBtn").disabled = false;
                        _("addRecordStatus").innerHTML = "";
                      } else {
                          _("addRecordBtn").disabled = false;
                          _("addRecordStatus").innerHTML = '<div style="color: red; margin-bottom: 10px;">Record already exist in this date.</div>';
                      }
                  }
                }
          ajax.send("date="+date+"&elco="+elco+"&elapp="+elapp+"&sponcom="+sponcom+"&gasera="+gasera+"&staele="+staele+"&lpg="+lpg+"&chem="+chem+"&phyro="+phyro+"&matlig="+matlig+"&light="+light+"&exp="+exp+"&resfire="+resfire+"&edufire="+edufire+"&hcfire="+hcfire+"&stofire="+stofire+"&bizfire="+bizfire+"&mixfire="+mixfire+"&indfire="+indfire+"&commer="+commer+"&assembly="+assembly+"&others="+others+"&grass="+grass+"&post="+post+"&veh="+veh+"&fatci="+fatci+"&fatbf="+fatbf+"&injci="+injci+"&injbf="+injbf);
          }
      }
      function updateRecord() {
      	var id = _("editRecordInputId").value;
      	_("updateRecordStatus").innerHTML = id;
      	// var date = _("add-date").value;
       //  var elco = _("add-el-co").value;
       //  var elapp = _("add-el-app").value;
       //  var sponcom = _("add-spon-com").value;
       //  var gasera = _("add-gasera").value;
       //  var staele = _("add-sta-ele").value;
       //  var lpg = _("add-lpg").value;
       //  var chem = _("add-chem").value;
       //  var phyro = _("add-phyro").value;
       //  var matlig = _("add-mat-lig").value;
       //  var light = _("add-light").value;
       //  var exp = _("add-exp").value;
       //  var resfire = _("add-res-fire").value;
       //  var edufire = _("add-edu-fire").value;
       //  var hcfire = _("add-hc-fire").value;
       //  var stofire = _("add-sto-fire").value;
       //  var bizfire = _("add-biz-fire").value;
       //  var mixfire = _("add-mix-fire").value;
       //  var indfire = _("add-ind-fire").value;
       //  var commer = _("add-com-mer").value;
       //  var assembly = _("add-assembly").value;
       //  var others = _("add-others").value;
       //  var grass = _("add-grass").value;
       //  var post = _("add-post").value;
       //  var veh = _("add-veh").value;
       //  var fatci = _("add-fat-ci").value;
       //  var fatbf = _("add-fat-bf").value;
       //  var injci = _("add-inj-ci").value;
       //  var injbf = _("add-inj-bf").value;
       //  var load = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i><center>';
       //  var error = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
       //  _("addRecordStatus").innerHTML = load;
       //    if (date == "" || elco == "" || elco == "0" || elapp == "" || elapp == "0" || sponcom == "" || sponcom == "0" || gasera == "" || gasera == "0" || staele == "" || staele == "0" || lpg == "" || lpg == "0" || chem == "" || chem == "0" || phyro == "" || phyro == "0" || matlig == "" || matlig == "0" || light == "" || light == "0" || exp == "" || exp == "0" || resfire == "" || resfire == "0" || edufire == "" || edufire == "0" || hcfire == "" || hcfire == "0" || stofire == "" || stofire == "0" || bizfire == "" || bizfire == "0" || mixfire == "" || mixfire == "0" || indfire == "" || indfire == "0" || commer == "" || commer == "0" || assembly == "" || assembly == "0" || others == "" || others == "0" || grass == "" || grass == "0" || post == "" || post == "0" || veh == "" || veh == "0" || fatci == "" || fatci == "0" || fatbf == "" || fatbf == "0" || injci == "" || injci == "0" || injbf == "" || injbf == "0") {
       //          _("addRecordStatus").innerHTML = error;
       //    } else {
       //          _("addRecordBtn").disabled = true;
       //          _("addRecordStatus").innerHTML = load;
       //          var ajax = ajaxObj("POST", "parsers/account.php");
       //          ajax.onreadystatechange = function() {
       //            if(ajaxReturn(ajax) == true) {
       //                if (ajax.responseText == "successinsert"){
       //                  _("add-date").value = "";
       //                	_("add-el-co").value = "";
       //                	_("add-el-app").value = "";
       //                	_("add-spon-com").value = "";
       //                	_("add-gasera").value = "";
       //                	_("add-sta-ele").value = "";
       //                	_("add-lpg").value = "";
       //                	_("add-chem").value = "";
       //                	_("add-phyro").value = "";
       //                	_("add-mat-lig").value = "";
       //                	_("add-light").value = "";
       //                	_("add-exp").value = "";
       //                	_("add-res-fire").value = "";
       //                	_("add-edu-fire").value = "";
       //                	_("add-hc-fire").value = "";
       //                	_("add-sto-fire").value = "";
       //                	_("add-biz-fire").value = "";
       //                	_("add-mix-fire").value = "";
       //                	_("add-ind-fire").value = "";
       //                	_("add-com-mer").value = "";
       //                	_("add-assembly").value = "";
       //                	_("add-others").value = "";
       //                	_("add-grass").value = "";
       //                	_("add-post").value = "";
       //                	_("add-veh").value = "";
       //                	_("add-fat-ci").value = "";
       //                	_("add-fat-bf").value = "";
       //                	_("add-inj-ci").value = "";
       //                	_("add-inj-bf").value = "";
       //                  var x = document.getElementById("snackbar-add-success")
       //                  x.className = "show";
       //                  setTimeout(function(){ x.className = x.className.replace("show", ""); }, 3000);
       //                  _("addRecordBtn").disabled = false;
       //                  _("addRecordStatus").innerHTML = "";
       //                } else {
       //                    _("addRecordBtn").disabled = false;
       //                    _("addRecordStatus").innerHTML = '<div style="color: red; margin-bottom: 10px;">Record already exist in this date.</div>';
       //                }
       //            }
       //          }
       //    ajax.send("date="+date+"&elco="+elco+"&elapp="+elapp+"&sponcom="+sponcom+"&gasera="+gasera+"&staele="+staele+"&lpg="+lpg+"&chem="+chem+"&phyro="+phyro+"&matlig="+matlig+"&light="+light+"&exp="+exp+"&resfire="+resfire+"&edufire="+edufire+"&hcfire="+hcfire+"&stofire="+stofire+"&bizfire="+bizfire+"&mixfire="+mixfire+"&indfire="+indfire+"&commer="+commer+"&assembly="+assembly+"&others="+others+"&grass="+grass+"&post="+post+"&veh="+veh+"&fatci="+fatci+"&fatbf="+fatbf+"&injci="+injci+"&injbf="+injbf);
       //    }
      }

      function addBarcodeMiddleText() {
        var code = _("add-barcode-middle-text").value;
        var load = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i><center>';
        var error = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
        var status = _("addBarcodeMiddleTextStatus");
        var button = _("addBarcodeMiddleTextBtn");
        status.innerHTML = load;
        if (code == "") {
            status.innerHTML = error;
        } else {
            button.disabled = true;
            status.innerHTML = load;
            var ajax = ajaxObj("POST", "parsers/account.php");
            ajax.onreadystatechange = function() {
              if(ajaxReturn(ajax) == true) {
	              if (ajax.responseText == "successinsert"){
	                  window.location = "account.php?id=<?php echo $id; ?>&masterlists=focus&select=barcode";
	              } else {
	                  button.disabled = false;
	                  status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Data already exist.</div>';
	              }
              }
            }
          ajax.send("addbarcodemiddletextcode="+code);
        }
      }

      function addDispatchControlMessenger() {
        var name = _("add-dispatch-control-messenger-name").value;
        var address = _("add-dispatch-control-messenger-address").value;
        var prepared = _("add-dispatch-control-messenger-prepared").value;
        var date = _("add-dispatch-control-messenger-date").value;
        var load = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i><center>';
        var error = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
        var status = _("addDispatchControlMessengerStatus");
        var button = _("addDispatchControlMessengerBtn");
        status.innerHTML = load;
        if (name == "" || address == "" || prepared == "" || date == "") {
            status.innerHTML = error;
        } else {
            button.disabled = true;
            status.innerHTML = load;
            var ajax = ajaxObj("POST", "parsers/account.php");
            ajax.onreadystatechange = function() {
              if(ajaxReturn(ajax) == true) {
	              if (ajax.responseText == "successinsert"){
	                  window.location = "account.php?id=<?php echo $id; ?>&dispatchcontrol=focus&select=view";
	              } else {
	                  button.disabled = false;
	                  status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Data already exist.</div>';
	              }
              }
            }
          ajax.send("adddispatchcontrolmessengername="+name+"&adddispatchcontrolmessengeraddress="+address+"&adddispatchcontrolmessengerprepared="+prepared+"&adddispatchcontrolmessengerdate="+date);
        }
      }
      
      function openBarcodeMiddleTextUpdateDialog(rid, code) {
     	$('#editBarcodeMiddleText').modal('show');
     	_("update-barcode-middle-text-id").value = rid;
     	_("update-barcode-middle-text-code").value = code;
      }

      function updateBarcodeMiddleTextRecord() {
      	var rid = _("update-barcode-middle-text-id").value;
        var code = _("update-barcode-middle-text-code").value;
        var load = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i><center>';
        var error = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
        var status = _("updateBarcodeMiddleTextStatus");
        var button = _("updateBarcodeMiddleTextBtn");
        status.innerHTML = load;
        if (code == "") {
            status.innerHTML = error;
        } else {
            button.disabled = true;
            status.innerHTML = load;
            var ajax = ajaxObj("POST", "parsers/account.php");
            ajax.onreadystatechange = function() {
              if(ajaxReturn(ajax) == true) {
	              if (ajax.responseText == "successupdate"){
	                  window.location = "account.php?id=<?php echo $id; ?>&masterlists=focus&select=barcode";
	              } else {
	                  button.disabled = false;
	                  status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Data already exist.</div>';
	              }
              }
            }
          ajax.send("updatebarcodemiddletextid="+rid+"&updatebarcodemiddletextcode="+code);
        }
      }

      function openAreaPriceUpdateDialog(rid, name, price) {
     	$('#editAreaPrice').modal('show');
     	_("update-area-price-id").value = rid;
     	_("update-area-price-name").value = name;
     	_("update-area-price-price").value = price;
      }

      function updateAreaPriceRecord() {
      	var rid = _("update-area-price-id").value;
        var price = _("update-area-price-price").value;
        var load = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i><center>';
        var error = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
        var status = _("updateAreaPriceStatus");
        var button = _("updateAreaPriceBtn");
        status.innerHTML = load;
        if (price == "") {
            status.innerHTML = error;
        } else {
            button.disabled = true;
            status.innerHTML = load;
            var ajax = ajaxObj("POST", "parsers/account.php");
            ajax.onreadystatechange = function() {
              if(ajaxReturn(ajax) == true) {
	              if (ajax.responseText == "successupdate"){
	                  window.location = "account.php?id=<?php echo $id; ?>&reportsettings=focus";
	              } else {
	                  button.disabled = false;
	                  status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Data already exist.</div>';
	              }
              }
            }
          ajax.send("updateareapriceid="+rid+"&updateareapriceprice="+price);
        }
      }

      function openBarcodeMiddleTextDeleteDialog(rid) {
     	$('#deleteBarcodeMiddleText').modal('show');
     	_("delete-barcode-middle-text-id").value = rid;
      }

      function deleteBarcodeMiddleTextRecord() {
      	var rid = _("delete-barcode-middle-text-id").value;
        var load = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i><center>';
        var error = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
        var status = _("deleteBarcodeMiddleTextStatus");
        var button = _("deleteBarcodeMiddleTextBtn");
        status.innerHTML = load;
        if (rid == "") {
            status.innerHTML = error;
        } else {
            button.disabled = true;
            status.innerHTML = load;
            var ajax = ajaxObj("POST", "parsers/account.php");
            ajax.onreadystatechange = function() {
              if(ajaxReturn(ajax) == true) {
	              if (ajax.responseText == "successupdate"){
	                  window.location = "account.php?id=<?php echo $id; ?>&masterlists=focus&select=barcode";
	              } else {
	                  button.disabled = false;
	                  status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Data already exist.</div>';
	              }
              }
            }
          ajax.send("deletebarcodemiddletextid="+rid);
        }
      }

      function openBarcodeMiddleTextDeleteDialog(rid) {
     	$('#deleteBarcodeMiddleText').modal('show');
     	_("delete-barcode-middle-text-id").value = rid;
      }

      function deleteBarcodeMiddleTextRecord() {
      	var rid = _("delete-barcode-middle-text-id").value;
        var load = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i><center>';
        var error = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
        var status = _("deleteBarcodeMiddleTextStatus");
        var button = _("deleteBarcodeMiddleTextBtn");
        status.innerHTML = load;
        if (rid == "") {
            status.innerHTML = error;
        } else {
            button.disabled = true;
            status.innerHTML = load;
            var ajax = ajaxObj("POST", "parsers/account.php");
            ajax.onreadystatechange = function() {
              if(ajaxReturn(ajax) == true) {
	              if (ajax.responseText == "successupdate"){
	                  window.location = "account.php?id=<?php echo $id; ?>&masterlists=focus&select=barcode";
	              } else {
	                  button.disabled = false;
	                  status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Unknown Error!</div>';
	              }
              }
            }
          ajax.send("deletebarcodemiddletextid="+rid);
        }
      }

      function openMasterListItemDeleteDialog(rid) {
     	$('#deleteMasterListItem').modal('show');
     	_("delete-master-list-item-id").value = rid;
      }

      function deleteMasterListItemRecord() {
      	var rid = _("delete-master-list-item-id").value;
      	var encodeid = "<?php if(isset($_GET["encodeid"])) { echo $_GET["encodeid"]; } ?>";
        var load = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i><center>';
        var error = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
        var status = _("deleteBarcodeMiddleTextStatus");
        var button = _("deleteBarcodeMiddleTextBtn");
        status.innerHTML = load;
        if (rid == "") {
            status.innerHTML = error;
        } else {
            button.disabled = true;
            status.innerHTML = load;
            var ajax = ajaxObj("POST", "parsers/account.php");
            ajax.onreadystatechange = function() {
              if(ajaxReturn(ajax) == true) {
	              if (ajax.responseText == "successupdate"){
	                  window.location = "account.php?id=<?php echo $id; ?>&masterlists=focus&select=view&encodeid="+encodeid;
	              } else {
	                  button.disabled = false;
	                  status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Unknown Error!</div>';
	              }
              }
            }
          ajax.send("deletemasterlistitemid="+rid);
        }
      }


      function openDispatchControlMessengerUpdateDialog(rid, name, address, prepared, date) {
     	$('#editDispatchControlMessenger').modal('show');
     	_("update-dispatch-control-messenger-id").value = rid;
     	_("update-dispatch-control-messenger-name").value = name;
     	_("update-dispatch-control-messenger-address").value = address;
     	_("update-dispatch-control-messenger-prepared").value = prepared;
     	_("update-dispatch-control-messenger-date").value = date;
      }

      function updateDispatchControlMessengerRecord() {
      	var rid = _("update-dispatch-control-messenger-id").value;
        var name = _("update-dispatch-control-messenger-name").value;
        var address = _("update-dispatch-control-messenger-address").value;
        var prepared = _("update-dispatch-control-messenger-prepared").value;
        var date = _("update-dispatch-control-messenger-date").value;
        var load = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i><center>';
        var error = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
        var status = _("updateDispatchControlMessengerStatus");
        var button = _("updateDispatchControlMessengerBtn");
        status.innerHTML = load;
        if (name == "" || address == "" || prepared == "" || date == "") {
            status.innerHTML = error;
        } else {
            button.disabled = true;
            status.innerHTML = load;
            var ajax = ajaxObj("POST", "parsers/account.php");
            ajax.onreadystatechange = function() {
              if(ajaxReturn(ajax) == true) {
	              if (ajax.responseText == "successupdate"){
	                  window.location = "account.php?id=<?php echo $id; ?>&dispatchcontrol=focus&select=view";
	              } else {
	                  button.disabled = false;
	                  status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Data already exist.</div>';
	              }
              }
            }
          ajax.send("updatedispatchcontrolmessengerid="+rid+"&updatedispatchcontrolmessengername="+name+"&updatedispatchcontrolmessengeaddress="+address+"&updatedispatchcontrolmessengerprepared="+prepared+"&updatedispatchcontrolmessengerdate="+date);
        }
      }

      function openDispatchControlMessengerDeleteDialog(rid) {
     	$('#deleteDispatchControlMessenger').modal('show');
     	_("delete-dispatch-control-messenger-id").value = rid;
      }

      function deleteDispatchControlMessengerRecord() {
      	var rid = _("delete-dispatch-control-messenger-id").value;
        var load = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i><center>';
        var error = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
        var status = _("deleteDispatchControlMessengerStatus");
        var button = _("deleteDispatchControlMessengerBtn");
        status.innerHTML = load;
        if (rid == "") {
            status.innerHTML = error;
        } else {
            button.disabled = true;
            status.innerHTML = load;
            var ajax = ajaxObj("POST", "parsers/account.php");
            ajax.onreadystatechange = function() {
              if(ajaxReturn(ajax) == true) {
	              if (ajax.responseText == "successupdate"){
	                  window.location = "account.php?id=<?php echo $id; ?>&dispatchcontrol=focus&select=view";
	              } else {
	                  button.disabled = false;
	                  status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Data already exist.</div>';
	              }
              }
            }
          ajax.send("deletedispatchcontrolmessengerid="+rid);
        }
      }

      function addDispatchControlDataRecord() {
        var rid = "<?php if(isset($_GET["recordidadd"])) { echo $_GET["recordidadd"]; } ?>";
        var cyclecode = _("add-dispatch-control-data-cycle-code").value;
        var pickupdate = _("add-dispatch-control-data-pickup-date").value;
        var sender = _("add-dispatch-control-data-sender").value;
        var deltype = _("add-dispatch-control-data-del-type").value;
        var load = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i><center>';
        var error = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
        var status = _("addDispatchControlDataStatus");
        var button = _("addDispatchControlDataBtn");
        status.innerHTML = load;
        if (cyclecode == "" || pickupdate == "" || sender == "" || deltype == "") {
            status.innerHTML = error;
        } else {
            button.disabled = true;
            status.innerHTML = load;
            var ajax = ajaxObj("POST", "parsers/account.php");
            ajax.onreadystatechange = function() {
              if(ajaxReturn(ajax) == true) {
	              if (ajax.responseText == "successinsert"){
	                  window.location = "account.php?id=<?php echo $id; ?>&dispatchcontrol=focus&select=view&recordidview="+rid;
	              } else if (ajax.responseText == "datanotexists"){
	                  button.disabled = false;
	                  status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Invalid Data! Make sure all of the data below are exist in the Master Lists single record.</div>';
	              } else if (ajax.responseText == "samedata"){
	                  button.disabled = false;
	                  status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Data already exist!</div>';
	              } else {
	                  button.disabled = false;
	                  status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Unknown Error!</div>';
	              }
              }
            }
          ajax.send("adddispatchcontroldatamessengerid="+rid+"&adddispatchcontroldatacyclecode="+cyclecode+"&adddispatchcontroldatapickupdate="+pickupdate+"&adddispatchcontroldatasender="+sender+"&adddispatchcontroldatadeltype="+deltype);
        }
      }

      function openDispatchControlDataViewDialog(rid, name, addres, prepared, date) {
     	$('#viewDispatchControlData').modal('show');
     	eraseCookie("add-dispatch-control-data-messenger-id");
     	createCookie("add-dispatch-control-data-messenger-id", rid, 1);
      }

      function openEncodeListsDeleteDialog(rid) {
     	$('#deleteEncodeLists').modal('show');
     	_("delete-encode-lists-id").value = rid;
      }

      function deleteEncodeListsRecord() {
      	var rid = _("delete-encode-lists-id").value;
        var load = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i><center>';
        var error = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
        var status = _("deleteEncodeListsStatus");
        var button = _("deleteEncodeListsBtn");
        status.innerHTML = load;
        if (rid == "") {
            status.innerHTML = error;
        } else {
            button.disabled = true;
            status.innerHTML = load;
            var ajax = ajaxObj("POST", "parsers/account.php");
            ajax.onreadystatechange = function() {
              if(ajaxReturn(ajax) == true) {
	              if (ajax.responseText == "successupdate"){
	                  window.location = "account.php?id=<?php echo $id; ?>&masterlists=focus&select=view";
	              } else {
	                  button.disabled = false;
	                  status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Data already exist.</div>';
	              }
              }
            }
          ajax.send("deleteencodelistsid="+rid);
        }
      }


       function encodeMasterListsRecord() {
      	var barcode = _("barcode-to-encode-search").value;
        var daterecieve = _("update-record-date-recieve").value;
        var receiveby = _("update-record-receive-by").value;
        var relation = _("update-record-relation").value;
        var messenger = _("update-record-messenger").value;
        var status = _("update-record-status").value;
        var reasonrts = _("update-record-reason-rts").value;
        var remarks = _("update-record-remarks").value;
        var datereported = _("update-record-date-reported").value;
        var load = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i><center>';
        var error = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
        var status = _("encodeMasterListsStatus");
        var button = _("encodeMasterListsBtn");
        status.innerHTML = load;
        if (barcode == "" || daterecieve == "" || receiveby == "" || relation == "" || messenger == "" || status == "" || reasonrts == "" || remarks == "" || datereported == "") {
            status.innerHTML = error;
        } else {
            button.disabled = true;
            status.innerHTML = load;
            var ajax = ajaxObj("POST", "parsers/account.php");
            ajax.onreadystatechange = function() {
              if(ajaxReturn(ajax) == true) {
	              if (ajax.responseText == "successupdate"){
	                  window.location = "account.php?id=<?php echo $id; ?>&encodemasterlists=focus";
	              } else {
	                  button.disabled = false;
	                  status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Data already exist.</div>';
	              }
              }
            }
          ajax.send("updaterecordbarcode="+barcode+"&updaterecorddaterecieve="+daterecieve+"&updaterecordrecieveby="+receiveby+"&updaterecordrelation="+relation+"&updaterecordmessenger="+messenger+"&updaterecordstatus="+status+"&updaterecordreasonrts="+reasonrts+"&updaterecordremarks="+remarks+"&updaterecorddatereported="+datereported);
        }
      }


	    $('form > .form-group > input').keyup(function() {

	        var empty = false;
	        $('form > .form-group > input').each(function() {
	            if ($(this).val() == '') {
	                empty = true;
	            }
	        });

	        if (empty) {
	            $('#uploadBtn').attr('disabled', 'disabled');
	        } else {
	            $('#uploadBtn').removeAttr('disabled');
	        }
	    });

	    $(function () {
	        $("#uploadFile, #import-barcode").bind("propertychange change click keyup input paste", function () {      
		      if ($("#uploadFile").val() != "" && $("#import-barcode").val() != "")
		          $('#uploadBtn').removeAttr('disabled');
		      else
		          $('#uploadBtn').attr('disabled', 'disabled');    
		      });
	    });

	    $(function () {
	        $("#importEncodeFile, #import-encode-sheet-number").bind("propertychange change click keyup input paste", function () {      
		      if ($("#importEncodeFile").val() != "" && $("#import-encode-sheet-number").val() != "")
		          $('#importEncodeBtn').removeAttr('disabled');
		      else
		          $('#importEncodeBtn').attr('disabled', 'disabled');    
		      });
	    });

	    $(function () {
	        $("#importAssignFile, #import-assign-sheet-number").bind("propertychange change click keyup input paste", function () {      
		      if ($("#importAssignFile").val() != "" && $("#import-assign-sheet-number").val() != "")
		          $('#importAssignBtn').removeAttr('disabled');
		      else
		          $('#importAssignBtn').attr('disabled', 'disabled');    
		      });
	    });

	    $(function () {
	        $("#summary-report-sender, #summary-report-deltype, #summary-report-month, #summary-report-year").bind("propertychange change keyup input paste", function () {      
		      if ($("#summary-report-sender").val() != "" && $("#summary-report-deltype").val() != "" && $("#summary-report-month").val() != "" && $("#summary-report-year").val() != "") {
		      	  var sender = $("#summary-report-sender").val();
		      	  var deltype = $("#summary-report-deltype").val();
		      	  var month = $("#summary-report-month").val();
		      	  var year = $("#summary-report-year").val();
		          $.ajax({  
	                url:"parsers/fetch-summary-pud-cyclecode.php",  
	                method:"post",  
	                data:{sender:sender, deltype:deltype, month:month, year:year},  
	                dataType:"text",  
	                success: function(data) {
			            $('#summary-report-checkbox-list').html(data); 
			        }

	              }); 
		      }
		      });
	    });

	    $(function () {
	        $("#days-report-sender, #days-report-deltype, #days-report-month, #days-report-year").bind("propertychange change keyup input paste", function () {      
		      if ($("#days-report-sender").val() != "" && $("#days-report-deltype").val() != "" && $("#days-report-month").val() != "" && $("#days-report-year").val() != "") {
		      	  var sender = $("#days-report-sender").val();
		      	  var deltype = $("#days-report-deltype").val();
		      	  var month = $("#days-report-month").val();
		      	  var year = $("#days-report-year").val();
		          $.ajax({  
	                url:"parsers/fetch-days-pud-cyclecode.php",  
	                method:"post",  
	                data:{sender:sender, deltype:deltype, month:month, year:year},  
	                dataType:"text",  
	                success: function(data) {
			            $('#days-report-checkbox-list').html(data); 
			        }

	              }); 
		      }
		      });
	    });

	    $(function () {
	        $("#report-sender, #report-deltype, #report-month, #report-year").bind("propertychange change keyup input paste", function () {      
		      if ($("#report-sender").val() != "" && $("#report-deltype").val() != "" && $("#report-month").val() != "" && $("#report-year").val() != "") {
		      	  var sender = $("#report-sender").val();
		      	  var deltype = $("#report-deltype").val();
		      	  var month = $("#report-month").val();
		      	  var year = $("#report-year").val();
		          $.ajax({  
	                url:"parsers/fetch-report-pud-cyclecode.php",  
	                method:"post",  
	                data:{sender:sender, deltype:deltype, month:month, year:year},  
	                dataType:"text",  
	                success: function(data) {
			            $('#report-checkbox-list').html(data); 
			        }

	              }); 
		      }
		      });
	    });


	    function request_page(pn){
	    	var encodeid = "<?php echo $encodeid; ?>";
	    	var txtSearch = _('record-search-text').value;
	           var txtFilter = _('record-filter-value').value;    
			 var results_box = _("record-search-result");

			 $.ajax({  
	                url:"parsers/fetch-search-count.php",  
	                method:"post",  
	                data:{search:txtSearch, filter:txtFilter, encodeid:encodeid},  
	                dataType:"text",  
	                success: function(data) {

			 var rpp = 10; // results per page
			 var last = data; // last page number
				var pagination_controls1 = _("pagination_controls1");
				var pagination_controls2 = _("pagination_controls2");
				results_box.innerHTML = "loading results ...";
			    var ajax = ajaxObj("POST", "parsers/fetch-master-lists.php");
	            ajax.onreadystatechange = function() {
	              if(ajax.readyState == 4 && ajax.status == 200) {
						var dataArray = ajax.responseText.split("||");
						var html_output = "";
						html_output += '<table class="table table-striped"><thead><tr><th>REC#</th><th>SENDER</th><th>DELTYPE</th><th>PUD</th><th>Month</th><th>JOB#</th><th>Checklist For PPB</th><th>File Name</th><th>seq no.</th><th>CYCLECODE</th><th>Qty</th><th>ADDRESS</th><th>AREA</th><th>SUBSCRIBERS NAME</th><th>BARCODE</th><th>ACCT NO</th><th>DATE RECEIVED</th><th>RECEIVED BY</th><th>RELATION</th><th>MESSENGER</th><th>STATUS</th><th>Reason for RTS</th><th>Remarks</th><th>Date Reported</th><th>Action</th></tr></thead><tbody>';
					    for(i = 0; i < dataArray.length - 1; i++){
							var itemArray = dataArray[i].split("|");
							html_output += '<tr><td>'+itemArray[0]+'</td><td>'+itemArray[1]+'</td><td>'+itemArray[2]+'</td><td>'+itemArray[3]+'</td><td>'+itemArray[4]+'</td><td>'+itemArray[5]+'</td><td>'+itemArray[6]+'</td><td>'+itemArray[7]+'</td><td>'+itemArray[8]+'</td><td>'+itemArray[9]+'</td><td>'+itemArray[10]+'</td><td>'+itemArray[11]+'</td><td>'+itemArray[12]+'</td><td>'+itemArray[13]+'</td><td>'+itemArray[14]+'</td><td>'+itemArray[15]+'</td><td>'+itemArray[16]+'</td><td>'+itemArray[17]+'</td><td>'+itemArray[18]+'</td><td>'+itemArray[19]+'</td><td>'+itemArray[20]+'</td><td>'+itemArray[21]+'</td><td>'+itemArray[22]+'</td><td>'+itemArray[23]+'</td><td><a href="account.php?id=1&masterlists=focus&select=view&encodeid='+encodeid+'&editrecid='+itemArray[0]+'">Edit</a> | <a href="#" onclick="openMasterListItemDeleteDialog(\''+itemArray[0]+'\')">Delete</a></td></tr>';
						}
						html_output += '</tbody></table>';
						results_box.innerHTML = html_output;
						console.log(ajax.responseText);
				    }
	            }
			    ajax.send("rpp="+rpp+"&last="+last+"&pn="+pn+"&search="+txtSearch+"&filter="+txtFilter+"&encodeid="+encodeid);
				// Change the pagination controls
				var paginationCtrls = "";
			    // Only if there is more than 1 page worth of results give the user pagination controls
			    if(last != 1){
					if (pn > 1) {
						paginationCtrls += '<button class="btn btn-default" id="minusBtn" onclick="request_page('+(pn-1)+')">&lt;</button>';
			    	}
					paginationCtrls += ' &nbsp; &nbsp; <b>Page '+pn+' of '+last+'</b> &nbsp; &nbsp; ';
			    	if (pn != last) {
			        	paginationCtrls += '<button class="btn btn-default" id="addBtn" onclick="request_page('+(pn+1)+')">&gt;</button>';
			    	}
			    }
				pagination_controls1.innerHTML = paginationCtrls;
				pagination_controls2.innerHTML = paginationCtrls;

				}

	            }); 
		}
		
	    $(document).ready(function(){  
	      $('#record-search-text').keyup(function(){ 
	      	   var encodeid = "<?php echo $encodeid; ?>";
	           var txtSearch = $(this).val();
	           var txtFilter = $('#record-filter-value').val();    
			   var pn = 1; // first page number

			             $.ajax({  
	                url:"parsers/fetch-search-count.php",  
	                method:"post",  
	                data:{search:txtSearch, filter:txtFilter, encodeid:encodeid},  
	                dataType:"text",  
	                success: function(data) {

	            var results_box = _("record-search-result");
	            var rpp = 10; // results per page
			 	var last = data; // last page number
				var pagination_controls1 = _("pagination_controls1");
				var pagination_controls2 = _("pagination_controls2");
				results_box.innerHTML = "loading results ...";
			    var ajax = ajaxObj("POST", "parsers/fetch-master-lists.php");
	            ajax.onreadystatechange = function() {
	              if(ajax.readyState == 4 && ajax.status == 200) {
						var dataArray = ajax.responseText.split("||");
						var html_output = "";
						html_output += '<table class="table table-striped"><thead><tr><th>REC#</th><th>SENDER</th><th>DELTYPE</th><th>PUD</th><th>Month</th><th>JOB#</th><th>Checklist For PPB</th><th>File Name</th><th>seq no.</th><th>CYCLECODE</th><th>Qty</th><th>ADDRESS</th><th>AREA</th><th>SUBSCRIBERS NAME</th><th>BARCODE</th><th>ACCT NO</th><th>DATE RECEIVED</th><th>RECEIVED BY</th><th>RELATION</th><th>MESSENGER</th><th>STATUS</th><th>Reason for RTS</th><th>Remarks</th><th>Date Reported</th><th>Action</th></tr></thead><tbody>';
					    for(i = 0; i < dataArray.length - 1; i++){
							var itemArray = dataArray[i].split("|");
							html_output += '<tr><td>'+itemArray[0]+'</td><td>'+itemArray[1]+'</td><td>'+itemArray[2]+'</td><td>'+itemArray[3]+'</td><td>'+itemArray[4]+'</td><td>'+itemArray[5]+'</td><td>'+itemArray[6]+'</td><td>'+itemArray[7]+'</td><td>'+itemArray[8]+'</td><td>'+itemArray[9]+'</td><td>'+itemArray[10]+'</td><td>'+itemArray[11]+'</td><td>'+itemArray[12]+'</td><td>'+itemArray[13]+'</td><td>'+itemArray[14]+'</td><td>'+itemArray[15]+'</td><td>'+itemArray[16]+'</td><td>'+itemArray[17]+'</td><td>'+itemArray[18]+'</td><td>'+itemArray[19]+'</td><td>'+itemArray[20]+'</td><td>'+itemArray[21]+'</td><td>'+itemArray[22]+'</td><td>'+itemArray[23]+'<td><a href="account.php?id=1&masterlists=focus&select=view&encodeid='+encodeid+'&editrecid='+itemArray[0]+'">Edit</a> | <a href="#" onclick="openMasterListItemDeleteDialog(\''+itemArray[0]+'\')">Delete</a></td></td></td></tr>';
						}
						html_output += '</tbody></table>';
						results_box.innerHTML = html_output;
						console.log(ajax.responseText);
				    }
	            }
			    ajax.send("rpp="+rpp+"&last="+last+"&pn="+pn+"&search="+txtSearch+"&filter="+txtFilter+"&encodeid="+encodeid);
				// Change the pagination controls
				var paginationCtrls = "";
			    // Only if there is more than 1 page worth of results give the user pagination controls
			    if(last != 1){
					if (pn > 1) {
						paginationCtrls += '<button class="btn btn-default" id="minusBtn" onclick="request_page('+(pn-1)+')">&lt;</button>';
			    	}
					paginationCtrls += ' &nbsp; &nbsp; <b>Page '+pn+' of '+last+'</b> &nbsp; &nbsp; ';
			    	if (pn != last) {
			        	paginationCtrls += '<button class="btn btn-default" id="addBtn" onclick="request_page('+(pn+1)+')">&gt;</button>';
			    	}
			    }
				pagination_controls1.innerHTML = paginationCtrls;
				pagination_controls2.innerHTML = paginationCtrls;

				}

	            });  

	      });  
	    });

	    function getLastPageNumber(keyword, filter) {
      	
          //   var ajax = ajaxObj("POST", "parsers/fetch-search-count.php");
          //   ajax.onreadystatechange = function() {
          //     if(ajaxReturn(ajax) == true) {
	         //      returbData = ajax.responseText;
          //     }
          //   }
          // ajax.send("search="+keyword+"&filter="+filter);

          $.ajax({  
	                url:"parsers/fetch-search-count.php",  
	                method:"post",  
	                data:{search:keyword, filter:filter},  
	                dataType:"text",  
	                success: function(data) {
			            return data;
			        }

	            });  
		        // search
        
      }


	    $(document).ready(function(){  
	      $('#dispatch-control-messenger-search-text').keyup(function(){  
	           var txtSearch = $(this).val(); 

	            $.ajax({  
	                url:"parsers/fetch-dispatch-control-messenger-search.php",  
	                method:"post",  
	                data:{search:txtSearch},  
	                dataType:"text",  
	                success:function(data) {  
	                    $('#dispatch-control-messenger-search-result').html(data);  
	                }  
	            });  

	      });  
	    });

	    $(document).ready(function(){  
	      $('#barcode-to-encode-search').bind("cut copy paste keyup",function(e){  
	           var txtSearch = $(this).val();

	            $.ajax({ 
	                url:"parsers/fetch-to-encode.php",  
	                method:"post",  
	                data:{check:"1",search:txtSearch},  
	                dataType:"text",  
	                success:function(data) {
	                	if($.trim(data) == "true") {
	                    	$('#to-encode-result').show();
	                    	$('#to-encode-result-no-data-found').hide();  
	                    } else if($.trim(data) == "false") {
	                    	$('#to-encode-result').hide();
							$('#to-encode-result-no-data-found').show(); 
	                    }
	                }  
	            });

	            $.ajax({ 
	                url:"parsers/fetch-to-encode.php",  
	                method:"post",  
	                data:{check:"2",search:txtSearch},  
	                dataType:"text",  
	                success:function(data) {
	                	$('#update-record-date-recieve').val(data); 
	                }  
	            });

	            $.ajax({ 
	                url:"parsers/fetch-to-encode.php",  
	                method:"post",  
	                data:{check:"3",search:txtSearch},  
	                dataType:"text",  
	                success:function(data) {
	                	$('#update-record-receive-by').val(data); 
	                }  
	            });    

	            $.ajax({ 
	                url:"parsers/fetch-to-encode.php",  
	                method:"post",  
	                data:{check:"4",search:txtSearch},  
	                dataType:"text",  
	                success:function(data) {
	                	$('#update-record-relation').val(data); 
	                }  
	            });    

	            $.ajax({ 
	                url:"parsers/fetch-to-encode.php",  
	                method:"post",  
	                data:{check:"5",search:txtSearch},  
	                dataType:"text",  
	                success:function(data) {
	                	$('#update-record-messenger').val(data); 
	                }  
	            });    

	            $.ajax({ 
	                url:"parsers/fetch-to-encode.php",  
	                method:"post",  
	                data:{check:"6",search:txtSearch},  
	                dataType:"text",  
	                success:function(data) {
	                	$('#update-record-status').val(data); 
	                }  
	            });    

	            $.ajax({ 
	                url:"parsers/fetch-to-encode.php",  
	                method:"post",  
	                data:{check:"7",search:txtSearch},  
	                dataType:"text",  
	                success:function(data) {
	                	$('#update-record-reason-rts').val(data); 
	                }  
	            });    

	            $.ajax({ 
	                url:"parsers/fetch-to-encode.php",  
	                method:"post",  
	                data:{check:"8",search:txtSearch},  
	                dataType:"text",  
	                success:function(data) {
	                	$('#update-record-remarks').val(data); 
	                }  
	            });    

	            $.ajax({ 
	                url:"parsers/fetch-to-encode.php",  
	                method:"post",  
	                data:{check:"9",search:txtSearch},  
	                dataType:"text",  
	                success:function(data) {
	                	$('#update-record-date-reported').val(data); 
	                }  
	            });        

	            // alert(txtSearch);

	            // // window.location = "account.php?id=<?php echo $id; ?>&encodemasterlists=focus&barcode="+txtSearch;

	      });  
	    });

	    $(document).ready(function(){  
	      $('#barcode-to-assign').bind("cut copy paste keyup",function(e){
	      	   var cyclecode = "<?php if(isset($_GET['cyclecode']) && isset($_GET['messengerid'])) echo $_GET['cyclecode']; ?>";
	      	   var messengerid = "<?php if(isset($_GET['cyclecode']) && isset($_GET['messengerid'])) echo $_GET['messengerid']; ?>";   
	           var txtSearch = $(this).val();

	            $.ajax({ 
	                url:"parsers/fetch-to-assign.php",  
	                method:"post",  
	                data:{cyclecode:cyclecode,messengerid:messengerid,search:txtSearch},  
	                dataType:"text",  
	                success:function(data) {
	                	if($.trim(data) == "successupdate") {
	                    	$('#barcode-to-assign-status').html('<div style="color: green; text-align: center;">Receipt Assigned!</div>'); 
	                    } else if($.trim(data) == "failedupdate") {
	                    	$('#barcode-to-assign-status').html('<div style="color: red; text-align: center;">Barcode does not exist!</div>');
	                    }
	                }  
	            });

	      });  
	    });


	      $('#search-encode-lists-date').change(function(e){  
	           var txtSearch = $(this).val();

	            $.ajax({ 
	                url:"parsers/search-encode-lists.php",  
	                method:"post",  
	                data:{search:txtSearch},  
	                dataType:"text",  
	                success:function(data) {
	                	$('#encode-lists-result').html(data); 
	                }  
	            });

	      });  


	    $("#record-filter-value").change(function () {
			var filterValue = $("#record-filter-value").val();

			if (filterValue == "all") {
				$("#text-search-div").show();
				$("#deltype-search-div").hide();
			} else if (filterValue == "cyclecode") {
				$("#text-search-div").show();
				$("#deltype-search-div").hide();
			} else if (filterValue == "deltype") {
				$("#text-search-div").hide();
				$("#deltype-search-div").show();
			} else if (filterValue == "sender") {
				$("#text-search-div").show();
				$("#deltype-search-div").hide();
			} else if (filterValue == "name") {
				$("#text-search-div").show();
				$("#deltype-search-div").hide();
			} else {
					
			}	

		});

	    $("#record-search-deltype").change(function () {
			var delType = $("#record-search-deltype").val();
			var txtFilter = $('#record-filter-value').val();  
			var encodeid = "<?php echo $encodeid; ?>";
			var pn = 1; 
		
					// $.ajax({  
		   //              url:"parsers/fetch-record-search.php",  
		   //              method:"post",  
		   //              data:{search:delType,filter:txtFilter},  
		   //              dataType:"text",  
		   //              success:function(data) {  
		   //                  $('#record-search-result').html(data);  
		   //              }  
		   //          });  

		        $.ajax({  
	                url:"parsers/fetch-search-count.php",  
	                method:"post",  
	                data:{search:delType, filter:txtFilter, encodeid:encodeid},  
	                dataType:"text",  
	                success: function(data) {

	            var results_box = _("record-search-result");
	            var rpp = 10; // results per page
			 	var last = data; // last page number
				var pagination_controls1 = _("pagination_controls1");
				var pagination_controls2 = _("pagination_controls2");
				results_box.innerHTML = "loading results ...";
			    var ajax = ajaxObj("POST", "parsers/fetch-master-lists.php");
	            ajax.onreadystatechange = function() {
	              if(ajax.readyState == 4 && ajax.status == 200) {
						var dataArray = ajax.responseText.split("||");
						var html_output = "";
						html_output += '<table class="table table-striped"><thead><tr><th>REC#</th><th>SENDER</th><th>DELTYPE</th><th>PUD</th><th>Month</th><th>JOB#</th><th>Checklist For PPB</th><th>File Name</th><th>seq no.</th><th>CYCLECODE</th><th>Qty</th><th>ADDRESS</th><th>AREA</th><th>SUBSCRIBERS NAME</th><th>BARCODE</th><th>ACCT NO</th><th>DATE RECEIVED</th><th>RECEIVED BY</th><th>RELATION</th><th>MESSENGER</th><th>STATUS</th><th>Reason for RTS</th><th>Remarks</th><th>Date Reported</th><th>Action</th></tr></thead><tbody>';
					    for(i = 0; i < dataArray.length - 1; i++){
							var itemArray = dataArray[i].split("|");
							html_output += '<tr><td>'+itemArray[0]+'</td><td>'+itemArray[1]+'</td><td>'+itemArray[2]+'</td><td>'+itemArray[3]+'</td><td>'+itemArray[4]+'</td><td>'+itemArray[5]+'</td><td>'+itemArray[6]+'</td><td>'+itemArray[7]+'</td><td>'+itemArray[8]+'</td><td>'+itemArray[9]+'</td><td>'+itemArray[10]+'</td><td>'+itemArray[11]+'</td><td>'+itemArray[12]+'</td><td>'+itemArray[13]+'</td><td>'+itemArray[14]+'</td><td>'+itemArray[15]+'</td><td>'+itemArray[16]+'</td><td>'+itemArray[17]+'</td><td>'+itemArray[18]+'</td><td>'+itemArray[19]+'</td><td>'+itemArray[20]+'</td><td>'+itemArray[21]+'</td><td>'+itemArray[22]+'</td><td>'+itemArray[23]+'</td><td><a href="account.php?id=1&masterlists=focus&select=view&encodeid='+encodeid+'&editrecid='+itemArray[0]+'">Edit</a> | <a href="#" onclick="openMasterListItemDeleteDialog(\''+itemArray[0]+'\')">Delete</a></td></td></tr>';
						}
						html_output += '</tbody></table>';
						results_box.innerHTML = html_output;
						console.log(ajax.responseText);
				    }
	            }
			    ajax.send("rpp="+rpp+"&last="+last+"&pn="+pn+"&search="+delType+"&filter="+txtFilter+"&encodeid="+encodeid);
				// Change the pagination controls
				var paginationCtrls = "";
			    // Only if there is more than 1 page worth of results give the user pagination controls
			    if(last != 1){
					if (pn > 1) {
						paginationCtrls += '<button class="btn btn-default" id="minusBtn" onclick="request_page('+(pn-1)+')">&lt;</button>';
			    	}
					paginationCtrls += ' &nbsp; &nbsp; <b>Page '+pn+' of '+last+'</b> &nbsp; &nbsp; ';
			    	if (pn != last) {
			        	paginationCtrls += '<button class="btn btn-default" id="addBtn" onclick="request_page('+(pn+1)+')">&gt;</button>';
			    	}
			    }
				pagination_controls1.innerHTML = paginationCtrls;
				pagination_controls2.innerHTML = paginationCtrls;

				}

	            });  

							
		});

    $('#edit-pud-date').datepicker({
	    format: 'mm/dd/yy'
	 });

    $('#edit-date-receive-date').datepicker({
	    format: 'yyyy-mm-dd'
	 });

    $('#edit-date-reported-date').datepicker({
	    format: 'yyyy-mm-dd'
	 });

    function goBack() { 
    	window.history.back(); 
    }


	</script>	
</body>

</html>
