<?php
include_once("includes/db_conn.php");
require_once 'library/PHPExcel/Classes/PHPExcel.php';
header('Content-Type: text/html; charset=ISO-8859-1');
ini_set('max_execution_time', 19200); //300 sedb_connds = 5 minutes
ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);

$uploadFilePath = 'uploads/sample.xlsx';

$excelReader = PHPExcel_IOFactory::createReaderForFile($uploadFilePath);
$excelObj = $excelReader->load($uploadFilePath);
$totalSheet = $excelObj->getSheetCount();
$importedCount = 0;

$excel = new PHPExcel();

//SHEET NAME

$excel -> setActiveSheetIndex(0) 
        -> setTitle('M-List');

//STYLE

$topHeaderStyleFontBlackBoldCalibri = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => '000000'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$topHeaderStyleFontBlackCalibri = array(
    'font'  => array(
        'bold'  => false,
        'color' => array('rgb' => '000000'),
        'size'  => 10,
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

//TABLE HEADER MASTERLISTS  

$excel -> getActiveSheet() 
        -> setCellValue('A1','REC#')
        -> getStyle('A1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('A1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('A1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('A1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('B1','SENDER')
        -> getStyle('B1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('B1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('B1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('B1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('C1','DELTYPE')
        -> getStyle('C1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('C1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('C1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('C1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('D1','PUD')
        -> getStyle('D1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('D1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('D1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('D1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('E1','Month')
        -> getStyle('E1')
        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('E1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('E1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> getActiveSheet() 
        -> setCellValue('F1',"JOB#")
        -> getStyle('F1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('F1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('F1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('F1')
        -> getAlignment()
        -> setWrapText(true);

$excel -> getActiveSheet() 
        -> getStyle('F1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('G1',"Checklist For PPB")
        -> getStyle('G1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('G1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('G1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('G1')
        -> getAlignment()
        -> setWrapText(true);

$excel -> getActiveSheet() 
        -> getStyle('G1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('H1',"File Name")
        -> getStyle('H1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('H1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('H1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('H1')
        -> getAlignment()
        -> setWrapText(true);

$excel -> getActiveSheet() 
        -> getStyle('H1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('I1',"seq no.")
        -> getStyle('I1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('I1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('I1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('I1')
        -> getAlignment()
        -> setWrapText(true);

$excel -> getActiveSheet() 
        -> getStyle('I1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('J1','CYCLECODE')
        -> getStyle('J1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('J1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('J1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('J1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('K1','Qty')
        -> getStyle('K1')
        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('K1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('K1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> getActiveSheet() 
        -> setCellValue('L1','ADDRESS')
        -> getStyle('L1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('L1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('L1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('L1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('M1','AREA')
        -> getStyle('M1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('M1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('M1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('M1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('N1','SUBSCRIBERS NAME')
        -> getStyle('N1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('N1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('N1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('N1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('O1','BARCODE')
        -> getStyle('O1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('O1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('O1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('O1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('P1','ACCT NO')
        -> getStyle('P1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('P1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('P1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('P1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('Q1','DATE RECEIVED')
        -> getStyle('Q1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('Q1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('Q1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('Q1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('R1','RECEIVED BY')
        -> getStyle('R1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('R1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('R1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('R1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('S1','RELATION')
        -> getStyle('S1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('S1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('S1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('S1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('T1','MESSENGER')
        -> getStyle('T1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('T1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('T1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('T1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('U1','STATUS')
        -> getStyle('U1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('U1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('U1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('U1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('V1','Reason for RTS')
        -> getStyle('V1')
        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('V1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('V1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> getActiveSheet() 
        -> setCellValue('W1','Remarks')
        -> getStyle('W1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('W1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('W1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('W1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> getActiveSheet() 
        -> setCellValue('X1','Date Reported')
        -> getStyle('X1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> getActiveSheet() 
        -> getStyle('X1')
        -> applyFromArray($centerText);

$excel -> getActiveSheet() 
        -> getStyle('X1')
        -> applyFromArray($border);

$excel -> getActiveSheet() 
        -> getStyle('X1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel->getActiveSheet()->getRowDimension('1')->setRowHeight(-1);
$excel->getActiveSheet()->getColumnDimension('A')->setWidth(9);
$excel->getActiveSheet()->getColumnDimension('B')->setWidth(9);
$excel->getActiveSheet()->getColumnDimension('C')->setWidth(16);
$excel->getActiveSheet()->getColumnDimension('D')->setWidth(9);
$excel->getActiveSheet()->getColumnDimension('E')->setWidth(11);
$excel->getActiveSheet()->getColumnDimension('F')->setWidth(10);
$excel->getActiveSheet()->getColumnDimension('G')->setWidth(13);
$excel->getActiveSheet()->getColumnDimension('H')->setWidth(13);
$excel->getActiveSheet()->getColumnDimension('I')->setWidth(10);
$excel->getActiveSheet()->getColumnDimension('J')->setWidth(12);
$excel->getActiveSheet()->getColumnDimension('K')->setWidth(9);
$excel->getActiveSheet()->getColumnDimension('L')->setWidth(70);
$excel->getActiveSheet()->getColumnDimension('M')->setWidth(12);
$excel->getActiveSheet()->getColumnDimension('N')->setWidth(40);
$excel->getActiveSheet()->getColumnDimension('O')->setWidth(20);
$excel->getActiveSheet()->getColumnDimension('P')->setWidth(13);
$excel->getActiveSheet()->getColumnDimension('Q')->setWidth(15);
$excel->getActiveSheet()->getColumnDimension('R')->setWidth(14);
$excel->getActiveSheet()->getColumnDimension('S')->setWidth(15);
$excel->getActiveSheet()->getColumnDimension('T')->setWidth(14);
$excel->getActiveSheet()->getColumnDimension('U')->setWidth(11);
$excel->getActiveSheet()->getColumnDimension('V')->setWidth(25);
$excel->getActiveSheet()->getColumnDimension('W')->setWidth(23);
$excel->getActiveSheet()->getColumnDimension('X')->setWidth(18);

if($totalSheet >= 1 && 1 > 0) {

	$success = false;
	$count = 0;
	$recnum = 0;

	$worksheet = $excelObj->getSheet(3);
	$lastRow = $worksheet->getHighestRow();
	$toInsert = array();
	for ($rowexcel = 1; $rowexcel <= $lastRow; $rowexcel++) {
		/* Check If sheet not emprt */
		$recnum++;

		$rec = $recnum;
		$sender = mysqli_real_escape_string($db_conn, $worksheet->getCell('B'.$rowexcel)->getValue());
		$deltype = mysqli_real_escape_string($db_conn, $worksheet->getCell('C'.$rowexcel)->getValue());
		$pud = mysqli_real_escape_string($db_conn, $worksheet->getCell('D'.$rowexcel)->getValue());
		$month = mysqli_real_escape_string($db_conn, $worksheet->getCell('E'.$rowexcel)->getValue());

		$jobnum = mysqli_real_escape_string($db_conn, $worksheet->getCell('F'.$rowexcel)->getValue());
		$checklist = mysqli_real_escape_string($db_conn, $worksheet->getCell('G'.$rowexcel)->getValue());
		$filenamerecord = mysqli_real_escape_string($db_conn, $worksheet->getCell('H'.$rowexcel)->getValue());
		$seqnum = mysqli_real_escape_string($db_conn, $worksheet->getCell('I'.$rowexcel)->getValue());

		$cyclecode = mysqli_real_escape_string($db_conn, $worksheet->getCell('J'.$rowexcel)->getValue());

		$qty = mysqli_real_escape_string($db_conn, $worksheet->getCell('K'.$rowexcel)->getValue());
		$address = mysqli_real_escape_string($db_conn, $worksheet->getCell('L'.$rowexcel)->getValue());
		$area = mysqli_real_escape_string($db_conn, $worksheet->getCell('M'.$rowexcel)->getValue());
		$subs = mysqli_real_escape_string($db_conn, $worksheet->getCell('N'.$rowexcel)->getValue());
		$barcode = mysqli_real_escape_string($db_conn, $worksheet->getCell('O'.$rowexcel)->getValue());
		$acctnum = mysqli_real_escape_string($db_conn, $worksheet->getCell('P'.$rowexcel)->getValue());
		$datereceive = mysqli_real_escape_string($db_conn, $worksheet->getCell('Q'.$rowexcel)->getValue());
		$receiveby = mysqli_real_escape_string($db_conn, $worksheet->getCell('R'.$rowexcel)->getValue());
		$relation = mysqli_real_escape_string($db_conn, $worksheet->getCell('S'.$rowexcel)->getValue());
		$messenger = mysqli_real_escape_string($db_conn, $worksheet->getCell('T'.$rowexcel)->getValue());
		$status = mysqli_real_escape_string($db_conn, $worksheet->getCell('U'.$rowexcel)->getValue()); 
		$reasonrts = mysqli_real_escape_string($db_conn, $worksheet->getCell('V'.$rowexcel)->getValue());
		$remarks = mysqli_real_escape_string($db_conn, $worksheet->getCell('W'.$rowexcel)->getValue());
		$datereported = mysqli_real_escape_string($db_conn, $worksheet->getCell('X'.$rowexcel)->getValue());

		if($datereceive == "0000-00-00 00:00:00" || $status == "RTS") {
			$datereceive = "";
		} else {
			$datereceive = date("m/d/Y", strtotime($datereceive));
		}

		if($datereported == "0000-00-00 00:00:00") {
			$datereported = "";
		} else {
			$datereported = date("m/d/Y", strtotime($datereported));
		}

		if($barcode != "BARCODE") {

			// $sql = "UPDATE ppbms_record SET record_date_receive='$newDate1', record_receive_by='$receiveby', record_relation='$relation', record_messenger='$messenger', record_status='$status', record_reason_rts='$reasonrts', record_remarks='$remarks', record_date_reported='$newDate2' WHERE record_bar_code='$barcode' AND record_status_status = '1' LIMIT 1";
			// $query = mysqli_query($db_conn, $sql);
			// $success = true;
			// $importedCount++;

			// DATA LOOP			

			$data=array($rec,$sender,$deltype,$pud,$month,$jobnum,$checklist,$filenamerecord,$seqnum,$cyclecode,$qty,$address,$area,$subs,$barcode,$acctnum,$datereceive,$receiveby,$relation,$messenger,$status,$reasonrts,$remarks,$datereported);
			array_push($toInsert, $data);

	        // $excel -> getActiveSheet() 
	        //         -> setCellValue('A'.$recnum,$rec)
	        //         -> getStyle('A'.$recnum)
	        // -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('A'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('A'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //                 -> setCellValue('B'.$recnum,$sender)
	        //                 -> getStyle('B'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('B'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('B'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //                 -> setCellValue('C'.$recnum,$deltype)
	        //                 -> getStyle('C'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('C'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('C'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //                 -> setCellValue('D'.$recnum,$pud)
	        //                 -> getStyle('D'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('D'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('D'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //                 -> setCellValue('E'.$recnum,$month)
	        //                 -> getStyle('E'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('E'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('E'.$recnum)
	        //                 -> getFill()
	        //         -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
	        //         -> getStartColor()
	        //         -> setRGB('ffff00');

	        // $excel -> getActiveSheet() 
	        //                 -> setCellValue('F'.$recnum,$jobnum)
	        //                 -> getStyle('F'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('F'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('F'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('F'.$recnum)
	        //         -> getAlignment()
	        //         -> setWrapText(true);

	        // $excel -> getActiveSheet() 
	        //                 -> setCellValue('G'.$recnum,$checklist)
	        //                 -> getStyle('G'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('G'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('G'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('G'.$recnum)
	        //         -> getAlignment()
	        //         -> setWrapText(true);

	        // $excel -> getActiveSheet() 
	        //                 -> setCellValue('H'.$recnum,$filenamerecord)
	        //                 -> getStyle('H'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('H'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('H'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('H'.$recnum)
	        //         -> getAlignment()
	        //         -> setWrapText(true);

	        // $excel -> getActiveSheet() 
	        //                 -> setCellValue('I'.$recnum,$seqnum)
	        //                 -> getStyle('I'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('I'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('I'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('I'.$recnum)
	        //         -> getAlignment()
	        //         -> setWrapText(true);

	        // $excel -> getActiveSheet() 
	        //         -> setCellValue('J'.$recnum,$cyclecode)
	        //         -> getStyle('J'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('J'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('J'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //                 -> setCellValue('K'.$recnum,$qty)
	        //                 -> getStyle('K'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //                 -> getStyle('K'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('K'.$recnum)
	        //         -> getFill()
	        //         -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
	        //         -> getStartColor()
	        //         -> setRGB('ffff00');

	        // $excel -> getActiveSheet() 
	        //         -> setCellValue('L'.$recnum,$address)
	        //         -> getStyle('L'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel->getActiveSheet()->getStyle('L'.$recnum)
	        //         ->getFont()
	        //         ->setSize(7);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('L'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //         -> setCellValue('M'.$recnum,$area)
	        //         -> getStyle('M'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('M'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('M'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //         -> setCellValue('N'.$recnum,$subs)
	        //         -> getStyle('N'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel->getActiveSheet()->getStyle('N'.$recnum)
	        //         ->getFont()
	        //         ->setSize(7);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('N'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //         -> setCellValue('O'.$recnum,$barcode)
	        //         -> getStyle('O'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('O'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('O'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //         -> setCellValue('P'.$recnum,$acctnum)
	        //         -> getStyle('P'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('P'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('P'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //         -> setCellValue('Q'.$recnum,$datereceive)
	        //         -> getStyle('Q'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('Q'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('Q'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //         -> setCellValue('R'.$recnum,$receiveby)
	        //         -> getStyle('R'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('R'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //         -> setCellValue('S'.$recnum,$relation)
	        //         -> getStyle('S'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('S'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //         -> setCellValue('T'.$recnum,$messenger)
	        //         -> getStyle('T'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('T'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //         -> setCellValue('U'.$recnum,$status)
	        //         -> getStyle('U'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('U'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('U'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //         -> setCellValue('V'.$recnum,$reasonrts)
	        //         -> getStyle('V'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('V'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('V'.$recnum)
	        //         -> getFill()
	        //         -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
	        //         -> getStartColor()
	        //         -> setRGB('ffff00');

	        // $excel -> getActiveSheet() 
	        //         -> setCellValue('W'.$recnum,$remarks)
	        //         -> getStyle('W'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('W'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('W'.$recnum)
	        //         -> applyFromArray($border);

	        // $excel -> getActiveSheet() 
	        //         -> setCellValue('X'.$recnum,$datereported)
	        //         -> getStyle('X'.$recnum)
	        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('X'.$recnum)
	        //         -> applyFromArray($centerText);

	        // $excel -> getActiveSheet() 
	        //         -> getStyle('X'.$recnum)
	        //         -> applyFromArray($border);

		}

	}

	$excel->getActiveSheet()->fromArray($toInsert, null, 'A2');

	header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	header('Content-Disposition: attachment; filename="Report.xlsx"');
	header('Cache-Control: max-age=0');
	$file = PHPExcel_IOFactory::createWriter($excel,'Excel2007');
	$file -> save('php://output');

	if($success == true) {
		echo '<h3>success</h3>!';
	} else {
		echo '<button type="button" onclick="goBack();">Back</button> <h3>No data found!</h3>';
	}

} else {
	echo '<button type="button" onclick="goBack();">Back</button> <h3>Invalid Sheet Number!</h3>';
}

?>