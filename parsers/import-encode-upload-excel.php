<title>PPBMS - Encode Master Lists</title>
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
	
if(isset($_POST["import-encode-submit-excel"]) && isset($_POST["import-encode-sheet-number"])){

	$sheetNumber = $_POST["import-encode-sheet-number"];
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
				$datereceiveCell = $worksheet->getCell('B' . $rowexcel);
				$datereceive = $datereceiveCell->getValue();
				if(PHPExcel_Shared_Date::isDateTime($datereceiveCell) && $datereceive != "") {
				     $datereceive = date($format = "m/d/y", PHPExcel_Shared_Date::ExcelToPHP($datereceive)); 
				     $datereceive = mysqli_real_escape_string($db_conn, $datereceive); 
				} else {
					$datereceive = mysqli_real_escape_string($db_conn, $worksheet->getCell('B'.$rowexcel)->getValue());
				}
				$receiveby = mysqli_real_escape_string($db_conn, $worksheet->getCell('C'.$rowexcel)->getValue());
				$relation = mysqli_real_escape_string($db_conn, $worksheet->getCell('D'.$rowexcel)->getValue());
				$messenger = mysqli_real_escape_string($db_conn, $worksheet->getCell('E'.$rowexcel)->getValue());
				$status = mysqli_real_escape_string($db_conn, $worksheet->getCell('F'.$rowexcel)->getValue());
				$reasonrts = mysqli_real_escape_string($db_conn, $worksheet->getCell('G'.$rowexcel)->getValue());
				$remarks = mysqli_real_escape_string($db_conn, $worksheet->getCell('H'.$rowexcel)->getValue());
				$datereportedCell = $worksheet->getCell('I' . $rowexcel);
				$datereported = $datereportedCell->getValue();
				if(PHPExcel_Shared_Date::isDateTime($datereportedCell) && $datereported != "") {
				    $datereported = date($format = "m/d/y", PHPExcel_Shared_Date::ExcelToPHP($datereported)); 
				    $datereported = mysqli_real_escape_string($db_conn, $datereported); 
				} else {
					$datereported = mysqli_real_escape_string($db_conn, $worksheet->getCell('I'.$rowexcel)->getValue());
				}
				$newDate1 = date("Y-m-d H:i:s", strtotime($datereceive));
  				$newDate2 = date("Y-m-d H:i:s", strtotime($datereported));

					if($barcode != "BARCODE") {

						$sql = "UPDATE ppbms_record SET record_date_receive='$newDate1', record_receive_by='$receiveby', record_relation='$relation', record_messenger='$messenger', record_status='$status', record_reason_rts='$reasonrts', record_remarks='$remarks', record_date_reported='$newDate2' WHERE record_bar_code='$barcode' AND record_status_status = '1' LIMIT 1";
						$query = mysqli_query($db_conn, $sql);
						$success = true;
						$importedCount++;

					}

			}

			if($success == true) {
				header("location: ../account.php?id=".$log_id."&importmasterlists=focus&status=success&importedcount=".$importedCount);
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