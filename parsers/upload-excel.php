<title>PPBMS - Import Master Lists</title>
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
	
if(isset($_POST["submit-excel"]) && isset($_POST["import-barcode"])){

	$barcodeMiddle = $_POST["import-barcode"];
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
		$currentYear = date("Y");

		$pathInfo = pathinfo($filename);

		$filenametochange = '../uploads/'.$pathInfo['filename'].'.csv';

		$lastid = 0;

		$sql_last = "SELECT id FROM ppbms_record ORDER BY id DESC LIMIT 1";
	  	$query_last = mysqli_query($db_conn, $sql_last);
		while($row_last = mysqli_fetch_array($query_last)) {  
				$lastid = $row_last["id"];
		}

		if($barcodeMiddle != "manual") {
			// Read the existing excel file
			// $inputFileType = PHPExcel_IOFactory::identify("../uploads/".$filename);
			// $objReader = PHPExcel_IOFactory::createReader($inputFileType);
			// $objPHPExcel = $objReader->load("../uploads/".$filename);

			// Update it's data
			// Set active sheet index to the first sheet, so Excel opens this as the first sheet
			$excelObj->setActiveSheetIndex(0);
			$totalRows = $excelObj->getActiveSheet()->getHighestRow();
			$senderheader = $excelObj->getActiveSheet()->getCell('B1')->getValue();
			$sender = mysqli_real_escape_string($db_conn, $excelObj->getActiveSheet()->getCell('B2')->getValue());
			$deltype = mysqli_real_escape_string($db_conn, $excelObj->getActiveSheet()->getCell('C2')->getValue());
			$pud = mysqli_real_escape_string($db_conn, $excelObj->getActiveSheet()->getCell('D2')->getValue());
			$pudCell = $excelObj->getActiveSheet()->getCell('D2');
				$pud = $pudCell->getValue();
			if(PHPExcel_Shared_Date::isDateTime($pudCell) && $pud != "") {
				     $pud = date($format = "n/j/y", PHPExcel_Shared_Date::ExcelToPHP($pud)); 
				     $pud = mysqli_real_escape_string($db_conn, $pud); 
				} else {
					$pud = mysqli_real_escape_string($db_conn, $excelObj->getActiveSheet()->getCell('D2')->getValue());
				}
			$cyclecode = mysqli_real_escape_string($db_conn, $excelObj->getActiveSheet()->getCell('J2')->getValue());

			$sql_cycle_code_count = "SELECT id FROM ppbms_record WHERE record_sender = '$sender' AND record_deltype = '$deltype' AND record_cycle_code = '$cyclecode' AND record_pud = '$pud' AND record_status_status = '1'";
	  			$query_cycle_code_count = mysqli_query($db_conn, $sql_cycle_code_count);
				$cycle_code_count = mysqli_num_rows($query_cycle_code_count);

				if($cycle_code_count > 0) {
					echo '<button type="button" onclick="goBack();">Back</button> <h3>Same sender, delivery type, cycle code and pick-up date found!</h3>';
				} else if($senderheader != "SENDER") {
							echo '<button type="button" onclick="goBack();">Back</button> <h3>It looks like sheet number 1 is not a masterlist.</h3>';
				} else {

					$sql_record = "INSERT INTO ppbms_encode_list (encode_file_name, encode_lists_date, encode_lists_status) VALUES ('$filename', NOW(), '1')";
					$query = mysqli_query($db_conn, $sql_record);
					$encode_lists_id = mysqli_insert_id($db_conn);		

			try {

					// Add column headers
					for ($i=2; $i<=$totalRows; $i++) {
					$excelObj->getActiveSheet()->setCellValue('O'.$i, $cyclecode.$barcodeMiddle.$lastid);
					$lastid++;
					}
								
					// Generate an updated excel file
					// Redirect output to a client’s web browser (Excel2007)
					$objWriter = PHPExcel_IOFactory::createWriter($excelObj, 'Excel2007');

				    $objWriter->save('../uploads/new'.$filename);
				    $objWriter = PHPExcel_IOFactory::createWriter($excelObj, 'CSV');
	    			$objWriter->save($filenametochange);

	    		$importquery = mysqli_query($db_conn, '
					    LOAD DATA LOCAL INFILE "'.$filenametochange.'"
					        INTO TABLE ppbms_record
					        CHARACTER SET UTF8
					        FIELDS TERMINATED by \',\'
					        ENCLOSED BY \'"\'
					        ESCAPED BY \'\b\'
					        LINES TERMINATED BY \'\r\n\'
					        IGNORE 1 LINES
					        (encode_lists_id, record_sender, record_deltype, record_pud, record_month, record_job_num, record_check_list, record_file_name, record_seq_num, record_cycle_code, record_qty, record_address, record_area, record_subs_name, record_bar_code, record_acct_num, record_date_receive, record_receive_by, record_relation, record_messenger, record_status, record_reason_rts, record_remarks, record_date_reported)
					        SET encode_lists_id = "'.$encode_lists_id.'",  record_year = "'.$currentYear.'", record_status_status = "1"
					')or die(mysqli_error());

	    		if($importquery) {
					header("location: ../account.php?id=".$log_id."&masterlists=focus&select=view");
				}

				} catch (Exception $e) {
				    echo 'ERROR: ', $e->getMessage();
				    die();
				}
			}
		} else {
			// Read the existing excel file
			// $inputFileType = PHPExcel_IOFactory::identify("../uploads/".$filename);
			// $objReader = PHPExcel_IOFactory::createReader($inputFileType);
			// $objPHPExcel = $objReader->load("../uploads/".$filename);

			// Update it's data
			// Set active sheet index to the first sheet, so Excel opens this as the first sheet
			$excelObj->setActiveSheetIndex(0);
			$totalRows = $excelObj->getActiveSheet()->getHighestRow();
			$senderheader = $excelObj->getActiveSheet()->getCell('B1')->getValue();
			$sender = mysqli_real_escape_string($db_conn, $excelObj->getActiveSheet()->getCell('B2')->getValue());
			$deltype = mysqli_real_escape_string($db_conn, $excelObj->getActiveSheet()->getCell('C2')->getValue());
			$pud = mysqli_real_escape_string($db_conn, $excelObj->getActiveSheet()->getCell('D2')->getValue());
			$pudCell = $excelObj->getActiveSheet()->getCell('D2');
				$pud = $pudCell->getValue();
			if(PHPExcel_Shared_Date::isDateTime($pudCell) && $pud != "") {
				     $pud = date($format = "n/j/y", PHPExcel_Shared_Date::ExcelToPHP($pud)); 
				     $pud = mysqli_real_escape_string($db_conn, $pud); 
				} else {
					$pud = mysqli_real_escape_string($db_conn, $excelObj->getActiveSheet()->getCell('D2')->getValue());
				}
			$cyclecode = mysqli_real_escape_string($db_conn, $excelObj->getActiveSheet()->getCell('J2')->getValue());
			$barcode = mysqli_real_escape_string($db_conn, $excelObj->getActiveSheet()->getCell('O2')->getValue());

			$sql_cycle_code_count = "SELECT id FROM ppbms_record WHERE record_sender = '$sender' AND record_deltype = '$deltype' AND record_cycle_code = '$cyclecode' AND record_pud = '$pud' AND record_status_status = '1'";
	  			$query_cycle_code_count = mysqli_query($db_conn, $sql_cycle_code_count);
				$cycle_code_count = mysqli_num_rows($query_cycle_code_count);

				if($cycle_code_count > 0) {
					echo '<button type="button" onclick="goBack();">Back</button> <h3>Same sender, delivery type, cycle code and pick-up date found!</h3>';
				} else if($senderheader != "SENDER") {
							echo '<button type="button" onclick="goBack();">Back</button> <h3>It looks like sheet number 1 is not a masterlist.</h3>';
				} else if($barcode == "") {
							echo '<button type="button" onclick="goBack();">Back</button> <h3>Oh man! This sheet does not have its own barcode.</h3>';
				} else {

					$sql_record = "INSERT INTO ppbms_encode_list (encode_file_name, encode_lists_date, encode_lists_status) VALUES ('$filename', NOW(), '1')";
					$query = mysqli_query($db_conn, $sql_record);
					$encode_lists_id = mysqli_insert_id($db_conn);		

			try {

				$objWriter = PHPExcel_IOFactory::createWriter($excelObj, 'CSV');
	    		$objWriter->save($filenametochange);

	    		$importquery = mysqli_query($db_conn, '
					    LOAD DATA LOCAL INFILE "'.$filenametochange.'"
					        INTO TABLE ppbms_record
					        CHARACTER SET UTF8
					        FIELDS TERMINATED by \',\'
					        ENCLOSED BY \'"\'
					        ESCAPED BY \'\b\'
					        LINES TERMINATED BY \'\r\n\'
					        IGNORE 1 LINES
					        (encode_lists_id, record_sender, record_deltype, record_pud, record_month, record_job_num, record_check_list, record_file_name, record_seq_num, record_cycle_code, record_qty, record_address, record_area, record_subs_name, record_bar_code, record_acct_num, record_date_receive, record_receive_by, record_relation, record_messenger, record_status, record_reason_rts, record_remarks, record_date_reported)
					        SET encode_lists_id = "'.$encode_lists_id.'",  record_year = "'.$currentYear.'", record_status_status = "1"
					')or die(mysqli_error());

	    		if($importquery) {
					header("location: ../account.php?id=".$log_id."&masterlists=focus&select=view");
				}

				} catch (Exception $e) {
				    echo 'ERROR: ', $e->getMessage();
				    die();
				}
			}
		}

		// mysqli_query($db_conn, '
		//     LOAD DATA LOCAL INFILE "'.$filenametochange.'"
		//         INTO TABLE ppbms_record
		//         FIELDS TERMINATED by \',\'
		//         ENCLOSED BY \'"\'
		//         LINES TERMINATED BY \'\n\'
		//         IGNORE 1 LINES
		//         (encode_lists_id, record_sender, record_deltype, record_pud, record_month, record_job_num, record_check_list, record_file_name, record_seq_num, record_cycle_code, record_qty, record_address, record_area, record_subs_name, record_bar_code, record_acct_num, record_date_receive, record_receive_by, record_relation, record_messenger, record_status, record_reason_rts, record_remarks, record_date_reported)
		//         SET encode_lists_id = "1",  record_year = "2017"
		// ')or die(mysqli_error());

		// $result2=mysqli_query($db_conn,"select count(*) count from ppbms_record");
		// $r2=mysqli_fetch_array($result2);
		// $count2=(int)$r2['count'];

		// $count=$count2-$count1;
		// if($count>0)
		// echo "Success";
		// echo "<b> total $count records have been added to the table $table </b> ";


	}

  mysqli_close($db_conn);

}

?>