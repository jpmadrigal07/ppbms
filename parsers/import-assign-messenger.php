<title>PPBMS - Assign Messenger</title>
<link rel="icon" type="image/png" href="../img/fav.png" />
<script type="text/javascript">
	function goBack() { 
    	window.history.back(); 
    }
</script>
<?php
include_once("../includes/loginstatus.php");
require_once '../library/PHPExcel/Classes/PHPExcel.php';
header('Content-Type: text/html; charset=ISO-8859-1');
ini_set('max_execution_time', 19200); //300 sedb_connds = 5 minutes
	
if(isset($_POST["import-assign-submit-excel"]) && isset($_POST["import-assign-sheet-number"]) && isset($_POST["import-assign-cycle-code"]) && isset($_POST["import-assign-messenger-id"])){

	$sheetNumber = $_POST["import-assign-sheet-number"];
	$cycleCode = $_POST["import-assign-cycle-code"];
	$messengerId = $_POST["import-assign-messenger-id"];
	$allowed =  array('xlsx','xls');
	$filename = $_FILES['file']['name'];
	$ext = pathinfo($filename, PATHINFO_EXTENSION);
	if(!in_array($ext,$allowed) ) {
	    echo '<h3>Invalid File! Make sure it is an excel file and try uploading again.</h3>';
	} else {

		ini_set('display_errors', 1);
		ini_set('display_startup_errors', 1);
		error_reporting(E_ALL);
		
		$uploadFilePath = '../uploads/'.basename($_FILES['file']['name']);
		move_uploaded_file($_FILES['file']['tmp_name'], $uploadFilePath);

		$excelReader = PHPExcel_IOFactory::createReaderForFile($uploadFilePath);
		$excelObj = $excelReader->load($uploadFilePath);
		$totalSheet = $excelObj->getSheetCount();
		$importedCount = 0;

		// foreach ($Reader->sheets() as $Index => $Name) {
		// 	echo 'Sheet #'.$Index.': '.$Name;
		// }

		if($totalSheet >= $sheetNumber && $sheetNumber > 0) {

			// echo $totalSheet.' | '.$sheetNumber;
			$success = false;
			$count = 0;

			$worksheet = $excelObj->getSheet($sheetNumber-1);
			$lastRow = $worksheet->getHighestRow();
			for ($rowexcel = 1; $rowexcel <= $lastRow; $rowexcel++) {
				/* Check If sheet not emprt */
				$barcode = mysqli_real_escape_string($db_conn, $worksheet->getCell('A'.$rowexcel)->getValue());

					if($barcode != "BARCODE") {

						$sql = "UPDATE ppbms_record SET messenger_id='$messengerId' WHERE record_bar_code = '$barcode' AND record_cycle_code = '$cycleCode' AND record_status_status = '1' LIMIT 1";
    					$query = mysqli_query($db_conn, $sql);
						$success = true;
						$importedCount++;
					}

			}

			if($success == true) {
				header("location: ../account.php?id=".$log_id."&dispatchcontrol=focus&select=view&cyclecode=".$cycleCode."&messengerid=".$messengerId."&method=import&status=success&importedcount=".$importedCount);
			} else {
				echo '<button type="button" onclick="goBack();">Back</button> <h3>No data found!</h3>';
			}

		} else {
			echo '<button type="button" onclick="goBack();">Back</button> <h3>Invalid Sheet Number!</h3>';
	    }

	}

  mysqli_close($db_conn);

  exit(); 

}

?>