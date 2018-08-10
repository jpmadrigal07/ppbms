<?php
include_once("includes/loginstatus.php");
require_once 'library/PHPExcel/Classes/PHPExcel.php';
$excel = new PHPExcel();

//SHEET NAME

$excel -> setActiveSheetIndex(0) 
		-> setTitle('Summary Report');

//STYLE

$topHeaderStyleFontBlackBoldCalibri = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => '000000'),
        'size'  => 11,
        'name'  => 'Calibri'
    ));

$topHeaderStyleFontBlackCalibri = array(
    'font'  => array(
        'bold'  => false,
        'color' => array('rgb' => '000000'),
        'size'  => 11,
        'name'  => 'Calibri'
    ));

$tableHeaderStyleFontWhiteBoldCalibri = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => 'FFFFFF'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$tableHeaderStyleFontWhiteCalibri = array(
    'font'  => array(
        'bold'  => false,
        'color' => array('rgb' => 'FFFFFF'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$tableHeaderStyleFontRedCalibri = array(
    'font'  => array(
        'bold'  => false,
        'color' => array('rgb' => 'ff0000'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$centerText = array(
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
        )
    );

$border = array(
  'borders' => array(
    'allborders' => array(
      'style' => PHPExcel_Style_Border::BORDER_THIN
    )
  )
);


//TOP HEADER

$sender = $_POST['sender'];
$deltype = $_POST['deltype'];
$month = $_POST['month'];
$year = $_POST['year'];

$excel -> setActiveSheetIndex(0) 
		-> mergeCells('A1:E1');

$excel -> setActiveSheetIndex(0) 
		-> setCellValue('A1','Summary Report: '.$sender.' '.$deltype) 
		-> getStyle('A1')
        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
		-> mergeCells('A2:E2');

$excel -> setActiveSheetIndex(0) 
		-> setCellValue('A2','Pick Of Month: '.$month.' '.$year) 
		-> getStyle('A2')
        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri);

//TABLE HEADER

$excel -> setActiveSheetIndex(0) 
		-> setCellValue('A4','CYCLECODE')
		-> getStyle('A4')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('A4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('A4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('A4')
		-> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
		-> setCellValue('B4','Date Of Pick-Up')
		-> getStyle('B4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('B4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('B4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('B4')
		-> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
		-> setCellValue('C4','Volume')
		-> getStyle('C4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('C4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('C4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('C4')
		-> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
		-> setCellValue('D4','Delivered')
		-> getStyle('D4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('D4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('D4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('D4')
		-> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
		-> setCellValue('E4','RTS')
		-> getStyle('E4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('E4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('E4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('E4')
		-> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
		-> setCellValue('F4',"Out Of Scope\n(NCR)")
		-> getStyle('F4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('F4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('F4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('F4')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('F4')
		-> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
		-> setCellValue('G4',"Out Of Scope\n(Provincial)")
		-> getStyle('G4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('G4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('G4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('G4')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('G4')
		-> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
		-> setCellValue('H4',"Delivered\n(NCR)")
		-> getStyle('H4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('H4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('H4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('H4')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('H4')
		-> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
		-> setCellValue('I4',"Delivered\n(Provincial)")
		-> getStyle('I4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('I4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('I4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('I4')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('I4')
		-> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
		-> setCellValue('J4','No Status')
		-> getStyle('J4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('J4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('J4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('J4')
		-> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
		-> setCellValue('K4','Balance')
		-> getStyle('K4')
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('K4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('K4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
		-> getStyle('K4')
		-> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel->setActiveSheetIndex(0)->getRowDimension('4')->setRowHeight(-1);
$excel->setActiveSheetIndex(0)->getColumnDimension('A')->setWidth(12);
$excel->setActiveSheetIndex(0)->getColumnDimension('B')->setWidth(14);
$excel->setActiveSheetIndex(0)->getColumnDimension('C')->setWidth(9);
$excel->setActiveSheetIndex(0)->getColumnDimension('D')->setWidth(11);
$excel->setActiveSheetIndex(0)->getColumnDimension('E')->setWidth(8);
$excel->setActiveSheetIndex(0)->getColumnDimension('F')->setWidth(13);
$excel->setActiveSheetIndex(0)->getColumnDimension('G')->setWidth(13);
$excel->setActiveSheetIndex(0)->getColumnDimension('H')->setWidth(13);
$excel->setActiveSheetIndex(0)->getColumnDimension('I')->setWidth(13);
$excel->setActiveSheetIndex(0)->getColumnDimension('J')->setWidth(11);
$excel->setActiveSheetIndex(0)->getColumnDimension('K')->setWidth(9);

// DATA LOOP

$countSummary = count($_POST['checkboxsummary']);
$countStart = 4;
$footerCount = $countSummary+4;
$totalCountNCR = 0;
$totalCountPROV = 0;
$totalVolume = 0;
$totalDelivered = 0;
$totalRTS = 0;
$totalOOSNCR = 0;
$totalOOSPROV = 0;
$totalDeliveredNCR = 0;
$totalDeliveredPROV = 0;
$totalBalance = 0;
for($i=0; $i<$countSummary; $i++) {
	$summaryArray = explode(',', $_POST['checkboxsummary'][$i]);
	$cyclecode = $summaryArray[0];
	$pud = $summaryArray[1];
	$countStart++;

	// COUNT NCR VOLUME START

	$sql_total_ncr = "SELECT * FROM ppbms_record WHERE (record_area = 'PARANAQUE' OR record_area = 'MARIKINA' OR record_area = 'MUNTINLUPA' OR record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
	$query_total_ncr = mysqli_query($db_conn, $sql_total_ncr);
	$count_total_ncr = mysqli_num_rows($query_total_ncr);

	$totalCountNCR = $totalCountNCR+$count_total_ncr;

	// COUNT NCR VOLUME END

	// COUNT PROVINCIAL VOLUME START

	$sql_total_prov = "SELECT * FROM ppbms_record WHERE (record_area = 'RIZAL' OR record_area = 'LAGUNA') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
	$query_total_prov = mysqli_query($db_conn, $sql_total_prov);
	$count_total_prov = mysqli_num_rows($query_total_prov);

	$totalCountPROV = $totalCountPROV+$count_total_prov;

	// COUNT PROVINCIAL VOLUME END

	$excel -> setActiveSheetIndex(0) 
		-> setCellValue('A'.$countStart, $cyclecode)
		-> getStyle('A'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('A'.$countStart)
	        -> applyFromArray($centerText);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('A'.$countStart)
	        -> applyFromArray($border);

	// DIVIDER

	$excel -> setActiveSheetIndex(0) 
		-> setCellValue('B'.$countStart, $pud)
		-> getStyle('B'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('B'.$countStart)
	        -> applyFromArray($centerText);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('B'.$countStart)
	        -> applyFromArray($border);

	// DIVIDER

	$sql_volume = "SELECT * FROM ppbms_record WHERE record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
	$query_volume = mysqli_query($db_conn, $sql_volume);
	$count_volume = mysqli_num_rows($query_volume);

	$totalVolume = $totalVolume+$count_volume;

	$excel -> setActiveSheetIndex(0) 
		-> setCellValue('C'.$countStart, $count_volume)
		-> getStyle('C'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('C'.$countStart)
	        -> applyFromArray($centerText);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('C'.$countStart)
	        -> applyFromArray($border);

	// DIVIDER

	$sql_delivered = "SELECT * FROM ppbms_record WHERE record_status = 'DELIVERED' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
	$query_delivered = mysqli_query($db_conn, $sql_delivered);
	$count_delivered = mysqli_num_rows($query_delivered);

	$totalDelivered = $totalDelivered+$count_delivered;

	$excel -> setActiveSheetIndex(0) 
		-> setCellValue('D'.$countStart,$count_delivered)
		-> getStyle('D'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('D'.$countStart)
	        -> applyFromArray($centerText);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('D'.$countStart)
	        -> applyFromArray($border);

	// DIVIDER

	$sql_rts = "SELECT * FROM ppbms_record WHERE record_status = 'RTS' AND record_reason_rts != 'OUT OF SCOPE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
	$query_rts = mysqli_query($db_conn, $sql_rts);
	$count_rts = mysqli_num_rows($query_rts);

	$totalRTS = $totalRTS+$count_rts;

	$excel -> setActiveSheetIndex(0) 
		-> setCellValue('E'.$countStart,$count_rts)
		-> getStyle('E'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('E'.$countStart)
	        -> applyFromArray($centerText);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('E'.$countStart)
	        -> applyFromArray($border);

	// DIVIDER

	$sql_oos_ncr = "SELECT * FROM ppbms_record WHERE record_status = 'RTS' AND record_reason_rts = 'OUT OF SCOPE' AND (record_area = 'PARANAQUE' OR record_area = 'MARIKINA' OR record_area = 'MUNTINLUPA' OR record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
	$query_oos_ncr = mysqli_query($db_conn, $sql_oos_ncr);
	$count_oos_ncr = mysqli_num_rows($query_oos_ncr);

	$totalOOSNCR = $totalOOSNCR+$count_oos_ncr;

	$excel -> setActiveSheetIndex(0) 
		-> setCellValue('F'.$countStart,$count_oos_ncr)
		-> getStyle('F'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('F'.$countStart)
	        -> applyFromArray($centerText);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('F'.$countStart)
	        -> applyFromArray($border);

	// DIVIDER

	$sql_oos_prov = "SELECT * FROM ppbms_record WHERE record_status = 'RTS' AND record_reason_rts = 'OUT OF SCOPE' AND (record_area = 'LAGUNA' OR record_area = 'RIZAL') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
	$query_oos_prov = mysqli_query($db_conn, $sql_oos_prov);
	$count_oos_prov = mysqli_num_rows($query_oos_prov);

	$totalOOSPROV = $totalOOSPROV+$count_oos_prov;

	$excel -> setActiveSheetIndex(0) 
		-> setCellValue('G'.$countStart,$count_oos_prov)
		-> getStyle('G'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('G'.$countStart)
	        -> applyFromArray($centerText);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('G'.$countStart)
	        -> applyFromArray($border);

	// DIVIDER

	$sql_delivered_ncr = "SELECT * FROM ppbms_record WHERE record_status = 'DELIVERED' AND (record_area = 'PARANAQUE' OR record_area = 'MARIKINA' OR record_area = 'MUNTINLUPA' OR record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
	$query_delivered_ncr = mysqli_query($db_conn, $sql_delivered_ncr);
	$count_delivered_ncr = mysqli_num_rows($query_delivered_ncr);

	$totalDeliveredNCR = $totalDeliveredNCR+$count_delivered_ncr;

	$excel -> setActiveSheetIndex(0) 
		-> setCellValue('H'.$countStart,$count_delivered_ncr)
		-> getStyle('H'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('H'.$countStart)
	        -> applyFromArray($centerText);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('H'.$countStart)
	        -> applyFromArray($border);

	// DIVIDER

	$sql_delivered_prov = "SELECT * FROM ppbms_record WHERE record_status = 'DELIVERED' AND (record_area = 'LAGUNA' OR record_area = 'RIZAL') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
	$query_delivered_prov = mysqli_query($db_conn, $sql_delivered_prov);
	$count_delivered_prov = mysqli_num_rows($query_delivered_prov);

	$totalDeliveredPROV = $totalDeliveredPROV+$count_delivered_prov;

	$excel -> setActiveSheetIndex(0) 
		-> setCellValue('I'.$countStart,$count_delivered_prov)
		-> getStyle('I'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('I'.$countStart)
	        -> applyFromArray($centerText);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('I'.$countStart)
	        -> applyFromArray($border);

	// DIVIDER

	$excel -> setActiveSheetIndex(0) 
		-> setCellValue('J'.$countStart,'0')
		-> getStyle('J'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('J'.$countStart)
	        -> applyFromArray($centerText);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('J'.$countStart)
	        -> applyFromArray($border);

	// DIVIDER

	$sql_balance = "SELECT * FROM ppbms_record WHERE record_status = '' AND messenger_id != '0' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
	$query_balance = mysqli_query($db_conn, $sql_balance);
	$count_balance = mysqli_num_rows($query_balance);

	$totalBalance = $totalBalance+$count_balance;

	$excel -> setActiveSheetIndex(0) 
			-> setCellValue('K'.$countStart, $count_balance)
			-> getStyle('K'.$countStart)
			-> applyFromArray($tableHeaderStyleFontRedCalibri);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('K'.$countStart)
			-> applyFromArray($centerText);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('K'.$countStart)
			-> applyFromArray($border);

	$excel -> setActiveSheetIndex(0) 
			-> getStyle('K'.$countStart)
			-> getFill()
			-> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
			-> getStartColor()
			-> setRGB('ffff00');


	// 4 BLANK SPACE START

	if($i == $countSummary-1) {
		for($j=0; $j<8; $j++){
			$countStart++;

			if($j < 4) {

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('A'.$countStart, '')
					-> getStyle('A'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('A'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('A'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('B'.$countStart, '')
					-> getStyle('B'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('B'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('B'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('C'.$countStart, '')
					-> getStyle('C'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('C'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('C'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('D'.$countStart, '')
					-> getStyle('D'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('D'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('D'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('E'.$countStart, '')
					-> getStyle('E'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('E'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('E'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('F'.$countStart, '')
					-> getStyle('F'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('F'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('F'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('G'.$countStart, '')
					-> getStyle('G'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('G'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('G'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('H'.$countStart, '')
					-> getStyle('H'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('H'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('H'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('I'.$countStart, '')
					-> getStyle('I'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('I'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('I'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('J'.$countStart, '')
					-> getStyle('J'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('J'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('J'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
						-> setCellValue('K'.$countStart,'')
						-> getStyle('K'.$countStart)
				        -> applyFromArray($tableHeaderStyleFontRedCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('K'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('K'.$countStart)
				        -> applyFromArray($border);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('K'.$countStart)
						-> getFill()
				        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
				        -> getStartColor()
				        -> setRGB('ffff00');

			} else if ($j == 4) {

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('A'.$countStart, 'NCR')
					-> getStyle('A'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('A'.$countStart)
				        -> applyFromArray($border);

				 $excel -> setActiveSheetIndex(0) 
						-> getStyle('A'.$countStart)
						-> getFill()
				        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
				        -> getStartColor()
				        -> setRGB('d9d9d9');

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('B'.$countStart, '')
					-> getStyle('B'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('B'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('B'.$countStart)
				        -> applyFromArray($border);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('B'.$countStart)
						-> getFill()
				        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
				        -> getStartColor()
				        -> setRGB('d9d9d9');

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('C'.$countStart, $totalCountNCR)
					-> getStyle('C'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('C'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('C'.$countStart)
				        -> applyFromArray($border);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('C'.$countStart)
						-> getFill()
				        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
				        -> getStartColor()
				        -> setRGB('d9d9d9');

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('D'.$countStart, '')
					-> getStyle('D'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('D'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('D'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('E'.$countStart, '')
					-> getStyle('E'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('E'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('E'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('F'.$countStart, '')
					-> getStyle('F'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('F'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('F'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('G'.$countStart, '')
					-> getStyle('G'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('G'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('G'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('H'.$countStart, '')
					-> getStyle('H'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('H'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('H'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('I'.$countStart, '')
					-> getStyle('I'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('I'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('I'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('J'.$countStart, '')
					-> getStyle('J'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('J'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('J'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
						-> setCellValue('K'.$countStart,'')
						-> getStyle('K'.$countStart)
				        -> applyFromArray($tableHeaderStyleFontRedCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('K'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('K'.$countStart)
				        -> applyFromArray($border);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('K'.$countStart)
						-> getFill()
				        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
				        -> getStartColor()
				        -> setRGB('ffff00');

			} else if ($j == 5) {

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('A'.$countStart, 'PROVINCIAL')
					-> getStyle('A'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('A'.$countStart)
				        -> applyFromArray($border);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('A'.$countStart)
						-> getFill()
				        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
				        -> getStartColor()
				        -> setRGB('d9d9d9');

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('B'.$countStart, '')
					-> getStyle('B'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('B'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('B'.$countStart)
				        -> applyFromArray($border);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('B'.$countStart)
						-> getFill()
				        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
				        -> getStartColor()
				        -> setRGB('d9d9d9');

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('C'.$countStart, $totalCountPROV)
					-> getStyle('C'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('C'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('C'.$countStart)
				        -> applyFromArray($border);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('C'.$countStart)
						-> getFill()
				        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
				        -> getStartColor()
				        -> setRGB('d9d9d9');

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('D'.$countStart, '')
					-> getStyle('D'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('D'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('D'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('E'.$countStart, '')
					-> getStyle('E'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('E'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('E'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('F'.$countStart, '')
					-> getStyle('F'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('F'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('F'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('G'.$countStart, '')
					-> getStyle('G'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('G'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('G'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('H'.$countStart, '')
					-> getStyle('H'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('H'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('H'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('I'.$countStart, '')
					-> getStyle('I'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('I'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('I'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('J'.$countStart, '')
					-> getStyle('J'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('J'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('J'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
						-> setCellValue('K'.$countStart,'')
						-> getStyle('K'.$countStart)
				        -> applyFromArray($tableHeaderStyleFontRedCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('K'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('K'.$countStart)
				        -> applyFromArray($border);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('K'.$countStart)
						-> getFill()
				        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
				        -> getStartColor()
				        -> setRGB('ffff00');
				
			} else if ($j == 6) {

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('A'.$countStart, 'Total')
					-> getStyle('A'.$countStart)
			        -> applyFromArray($tableHeaderStyleFontRedCalibri);

			    $excel -> setActiveSheetIndex(0) 
						-> getStyle('A'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('A'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('B'.$countStart, '')
					-> getStyle('B'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('B'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('B'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('C'.$countStart, $totalCountNCR+$totalCountPROV)
					-> getStyle('C'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('C'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('C'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('D'.$countStart, $totalDelivered)
					-> getStyle('D'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('D'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('D'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('E'.$countStart, $totalRTS)
					-> getStyle('E'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('E'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('E'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('F'.$countStart, $totalOOSNCR)
					-> getStyle('F'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('F'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('F'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('G'.$countStart, $totalOOSPROV)
					-> getStyle('G'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('G'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('G'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('H'.$countStart, $totalDeliveredNCR)
					-> getStyle('H'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('H'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('H'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('I'.$countStart, $totalDeliveredPROV)
					-> getStyle('I'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('I'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('I'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('J'.$countStart, '0')
					-> getStyle('J'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('J'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('J'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
						-> setCellValue('K'.$countStart, $totalBalance)
						-> getStyle('K'.$countStart)
				        -> applyFromArray($tableHeaderStyleFontRedCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('K'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('K'.$countStart)
				        -> applyFromArray($border);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('K'.$countStart)
						-> getFill()
				        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
				        -> getStartColor()
				        -> setRGB('ffff00');
				
			} else if ($j == 7) {

				// PERCENTAGE START

				$totalCount = $totalCountNCR+$totalCountPROV;

				$delivered = ($totalDelivered/$totalCount)*100;
				$delivered = number_format($delivered, 2);

				$rts = ($totalRTS/$totalCount)*100;
				$rts = number_format($rts, 2);

				$oosncr = ($totalOOSNCR/$totalCount)*100;
				$oosncr = number_format($oosncr, 2);

				$oosprov = ($totalOOSPROV/$totalCount)*100;
				$oosprov = number_format($oosprov, 2);

				$deliveredncr = ($totalDeliveredNCR/$totalCount)*100;
				$deliveredncr = number_format($deliveredncr, 2);

				$deliveredprov = ($totalDeliveredPROV/$totalCount)*100;
				$deliveredprov = number_format($deliveredprov, 2);

				// PERCENTAGE END


				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('A'.$countStart, 'Percentage')
					-> getStyle('A'.$countStart)
			        -> applyFromArray($tableHeaderStyleFontRedCalibri);

			    $excel -> setActiveSheetIndex(0) 
						-> getStyle('A'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('A'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('B'.$countStart, '')
					-> getStyle('B'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('B'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('B'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('C'.$countStart, '')
					-> getStyle('C'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('C'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('C'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('D'.$countStart, $delivered.'%')
					-> getStyle('D'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('D'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('D'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('E'.$countStart, $rts.'%')
					-> getStyle('E'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('E'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('E'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('F'.$countStart, $oosncr.'%')
					-> getStyle('F'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('F'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('F'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('G'.$countStart, $oosprov.'%')
					-> getStyle('G'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('G'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('G'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('H'.$countStart, $deliveredncr.'%')
					-> getStyle('H'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('H'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('H'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('I'.$countStart, $deliveredprov.'%')
					-> getStyle('I'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('I'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('I'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
					-> setCellValue('J'.$countStart, '0.00%')
					-> getStyle('J'.$countStart)
			        -> applyFromArray($topHeaderStyleFontBlackCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('J'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('J'.$countStart)
				        -> applyFromArray($border);

				// DIVIDER

				$excel -> setActiveSheetIndex(0) 
						-> setCellValue('K'.$countStart,'')
						-> getStyle('K'.$countStart)
				        -> applyFromArray($tableHeaderStyleFontRedCalibri);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('K'.$countStart)
				        -> applyFromArray($centerText);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('K'.$countStart)
				        -> applyFromArray($border);

				$excel -> setActiveSheetIndex(0) 
						-> getStyle('K'.$countStart)
						-> getFill()
				        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
				        -> getStartColor()
				        -> setRGB('ffff00');
				
			}

		}
	}

	// 4 BLANK SPACE END

}

$deltype = preg_replace('/\s+/', '', $deltype);
$sender = preg_replace('/\s+/', '', $sender);

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="SummaryReport'.$sender.''.$deltype.''.$month.''.$year.'.xlsx"');
header('Cache-Control: max-age=0');
$file = PHPExcel_IOFactory::createWriter($excel,'Excel2007');
$file -> save('php://output');

mysqli_close($db_conn);

?>

