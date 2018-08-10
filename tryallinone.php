<?php
include_once("includes/loginstatus.php");
require_once 'library/PHPExcel/Classes/PHPExcel.php';
header('Content-Type: text/html; charset=ISO-8859-1');
ini_set('max_execution_time', 19200); //300 sedb_connds = 5 minutes
ini_set('memory_limit', '-1');


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

// POST

$sender = $_POST['sender'];
$deltype = $_POST['deltype'];
$month = $_POST['month'];
$year = $_POST['year'];

$sql_area_price = "SELECT area_price_price FROM ppbms_area_price WHERE area_price_status='1'";
$query_area_price = mysqli_query($db_conn, $sql_area_price);
$apcount = 0;
$pqueprice = 0;
$lpnasprice = 0;
$mlupaprice = 0;
$mknaprice = 0;
$rizalprice = 0;
$lagunaprice = 0;

while ($row = mysqli_fetch_array($query_area_price, MYSQLI_ASSOC)) {
        $apcount++;
        $areaprice = $row["area_price_price"];
        if($apcount == 1) {
            $pqueprice = $areaprice;
        } else if($apcount == 2) {
            $lpnasprice = $areaprice;
        } else if($apcount == 3) {
            $mlupaprice = $areaprice;
        } else if($apcount == 4) {
            $mknaprice = $areaprice;
        } else if($apcount == 5) {
            $rizalprice = $areaprice;
        } else if($apcount == 6) {
            $lagunaprice = $areaprice;
        }
}


//TABLE HEADER MASTERLISTS  

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('A1','REC#')
        -> getStyle('A1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('B1','SENDER')
        -> getStyle('B1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('C1','DELTYPE')
        -> getStyle('C1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('D1','PUD')
        -> getStyle('D1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('D1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('D1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('D1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('E1','Month')
        -> getStyle('E1')
        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('E1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('E1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('F1',"JOB#")
        -> getStyle('F1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F1')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('G1',"Checklist For PPB")
        -> getStyle('G1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G1')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('H1',"File Name")
        -> getStyle('H1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H1')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('I1',"seq no.")
        -> getStyle('I1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I1')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('J1','CYCLECODE')
        -> getStyle('J1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('K1','Qty')
        -> getStyle('K1')
        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('L1','ADDRESS')
        -> getStyle('L1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('M1','AREA')
        -> getStyle('M1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('N1','SUBSCRIBERS NAME')
        -> getStyle('N1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('O1','BARCODE')
        -> getStyle('O1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('P1','ACCT NO')
        -> getStyle('P1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Q1','DATE RECEIVED')
        -> getStyle('Q1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('R1','RECEIVED BY')
        -> getStyle('R1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('S1','RELATION')
        -> getStyle('S1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('T1','MESSENGER')
        -> getStyle('T1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('U1','STATUS')
        -> getStyle('U1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('V1','Reason for RTS')
        -> getStyle('V1')
        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('W1','Remarks')
        -> getStyle('W1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('X1','Date Reported')
        -> getStyle('X1')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X1')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X1')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel->setActiveSheetIndex(0)->getRowDimension('1')->setRowHeight(-1);
$excel->setActiveSheetIndex(0)->getColumnDimension('A')->setWidth(9);
$excel->setActiveSheetIndex(0)->getColumnDimension('B')->setWidth(9);
$excel->setActiveSheetIndex(0)->getColumnDimension('C')->setWidth(16);
$excel->setActiveSheetIndex(0)->getColumnDimension('D')->setWidth(9);
$excel->setActiveSheetIndex(0)->getColumnDimension('E')->setWidth(11);
$excel->setActiveSheetIndex(0)->getColumnDimension('F')->setWidth(10);
$excel->setActiveSheetIndex(0)->getColumnDimension('G')->setWidth(13);
$excel->setActiveSheetIndex(0)->getColumnDimension('H')->setWidth(13);
$excel->setActiveSheetIndex(0)->getColumnDimension('I')->setWidth(10);
$excel->setActiveSheetIndex(0)->getColumnDimension('J')->setWidth(12);
$excel->setActiveSheetIndex(0)->getColumnDimension('K')->setWidth(9);
$excel->setActiveSheetIndex(0)->getColumnDimension('L')->setWidth(70);
$excel->setActiveSheetIndex(0)->getColumnDimension('M')->setWidth(12);
$excel->setActiveSheetIndex(0)->getColumnDimension('N')->setWidth(40);
$excel->setActiveSheetIndex(0)->getColumnDimension('O')->setWidth(20);
$excel->setActiveSheetIndex(0)->getColumnDimension('P')->setWidth(13);
$excel->setActiveSheetIndex(0)->getColumnDimension('Q')->setWidth(15);
$excel->setActiveSheetIndex(0)->getColumnDimension('R')->setWidth(14);
$excel->setActiveSheetIndex(0)->getColumnDimension('S')->setWidth(15);
$excel->setActiveSheetIndex(0)->getColumnDimension('T')->setWidth(14);
$excel->setActiveSheetIndex(0)->getColumnDimension('U')->setWidth(11);
$excel->setActiveSheetIndex(0)->getColumnDimension('V')->setWidth(25);
$excel->setActiveSheetIndex(0)->getColumnDimension('W')->setWidth(23);
$excel->setActiveSheetIndex(0)->getColumnDimension('X')->setWidth(18);

// DATA LOOP

$countSummary = count($_POST['checkboxsummary']);

$count0 = 1;

$toInsert = array();

$data=array();

for($i=0; $i<$countSummary; $i++) {
    $summaryArray = explode(',', $_POST['checkboxsummary'][$i]);
    $cyclecode = $summaryArray[0];
    $pud = $summaryArray[1];
    $recnum = 0;

    $sql_master_lists = "SELECT record_sender, record_deltype, record_pud, record_month, record_job_num, record_check_list, record_file_name, record_seq_num, record_cycle_code, record_qty, record_address, record_area, record_subs_name, record_bar_code, record_acct_num, record_date_receive, record_receive_by, record_relation, record_messenger, record_status, record_reason_rts, record_remarks, record_date_reported FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_master_lists = mysqli_query($db_conn, $sql_master_lists);

    while ($row = mysqli_fetch_array($query_master_lists, MYSQLI_ASSOC)) {
          $count0++;
          $recnum++;

          $rec = $recnum;
          $sender = isset($row["record_sender"]) ? $row['record_sender'] : '';
          $deltype = isset($row["record_deltype"]) ? $row['record_deltype'] : ''; 
          $pud = isset($row["record_pud"]) ? $row['record_pud'] : ''; 
          $month = isset($row["record_month"]) ? $row['record_month'] : '';

          $jobnum = isset($row["record_job_num"]) ? $row['record_job_num'] : '';
          $checklist = isset($row["record_check_list"]) ? $row['record_check_list'] : ''; 
          $filenamerecord = isset($row["record_file_name"]) ? $row['record_file_name'] : '';
          $seqnum = isset($row["record_seq_num"]) ? $row['record_seq_num'] : '';

          $cyclecode = isset($row["record_cycle_code"]) ? $row['record_cycle_code'] : '';

          $qty = isset($row["record_qty"]) ? $row['record_qty'] : '';
          $address = isset($row["record_address"]) ? $row['record_address'] : '';
          $area = isset($row["record_area"]) ? $row['record_area'] : '';
          $subs = isset($row["record_subs_name"]) ? $row['record_subs_name'] : '';
          $barcode = isset($row["record_bar_code"]) ? $row['record_bar_code'] : '';
          $acctnum = isset($row["record_acct_num"]) ? $row['record_acct_num'] : '';
          $datereceive = isset($row["record_date_receive"]) ? $row['record_date_receive'] : '';
          $receiveby = isset($row["record_receive_by"]) ? $row['record_receive_by'] : ''; 
          $relation = isset($row["record_relation"]) ? $row['record_relation'] : ''; 
          $messenger = isset($row["record_messenger"]) ? $row['record_messenger'] : ''; 
          $status = isset($row["record_status"]) ? $row['record_status'] : ''; 
          $reasonrts = isset($row["record_reason_rts"]) ? $row['record_reason_rts'] : ''; 
          $remarks = isset($row["record_remarks"]) ? $row['record_remarks'] : '';
          $datereported = isset($row["record_date_reported"]) ? $row['record_date_reported'] : '';

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

          $data=array($rec,$sender,$deltype,$pud,$month,$jobnum,$checklist,$filenamerecord,$seqnum,$cyclecode,$qty,$address,$area,$subs,$barcode,$acctnum,$datereceive,$receiveby,$relation,$messenger,$status,$reasonrts,$remarks,$datereported);
          array_push($toInsert, $data);

        // $excel -> setActiveSheetIndex(0) 
        //         -> setCellValue('A'.$count0,$rec)
        //         -> getStyle('A'.$count0)
        // -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('A'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('A'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> setCellValue('B'.$count0,$sender)
        //                 -> getStyle('B'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('B'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('B'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> setCellValue('C'.$count0,$deltype)
        //                 -> getStyle('C'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('C'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('C'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> setCellValue('D'.$count0,$pud)
        //                 -> getStyle('D'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('D'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('D'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> setCellValue('E'.$count0,$month)
        //                 -> getStyle('E'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('E'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('E'.$count0)
        //                 -> getFill()
        //         -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        //         -> getStartColor()
        //         -> setRGB('ffff00');

        // $excel -> setActiveSheetIndex(0) 
        //                 -> setCellValue('F'.$count0,$jobnum)
        //                 -> getStyle('F'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('F'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('F'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('F'.$count0)
        //         -> getAlignment()
        //         -> setWrapText(true);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> setCellValue('G'.$count0,$checklist)
        //                 -> getStyle('G'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('G'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('G'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('G'.$count0)
        //         -> getAlignment()
        //         -> setWrapText(true);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> setCellValue('H'.$count0,$filenamerecord)
        //                 -> getStyle('H'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('H'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('H'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('H'.$count0)
        //         -> getAlignment()
        //         -> setWrapText(true);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> setCellValue('I'.$count0,$seqnum)
        //                 -> getStyle('I'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('I'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('I'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('I'.$count0)
        //         -> getAlignment()
        //         -> setWrapText(true);

        // $excel -> setActiveSheetIndex(0) 
        //         -> setCellValue('J'.$count0,$cyclecode)
        //         -> getStyle('J'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('J'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('J'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> setCellValue('K'.$count0,$qty)
        //                 -> getStyle('K'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //                 -> getStyle('K'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('K'.$count0)
        //         -> getFill()
        //         -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        //         -> getStartColor()
        //         -> setRGB('ffff00');

        // $excel -> setActiveSheetIndex(0) 
        //         -> setCellValue('L'.$count0,$address)
        //         -> getStyle('L'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel->getActiveSheet()->getStyle('L'.$count0)
        //         ->getFont()
        //         ->setSize(7);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('L'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //         -> setCellValue('M'.$count0,$area)
        //         -> getStyle('M'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('M'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('M'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //         -> setCellValue('N'.$count0,$subs)
        //         -> getStyle('N'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel->getActiveSheet()->getStyle('N'.$count0)
        //         ->getFont()
        //         ->setSize(7);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('N'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //         -> setCellValue('O'.$count0,$barcode)
        //         -> getStyle('O'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('O'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('O'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //         -> setCellValue('P'.$count0,$acctnum)
        //         -> getStyle('P'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('P'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('P'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //         -> setCellValue('Q'.$count0,$datereceive)
        //         -> getStyle('Q'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('Q'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('Q'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //         -> setCellValue('R'.$count0,$receiveby)
        //         -> getStyle('R'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('R'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //         -> setCellValue('S'.$count0,$relation)
        //         -> getStyle('S'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('S'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //         -> setCellValue('T'.$count0,$messenger)
        //         -> getStyle('T'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('T'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //         -> setCellValue('U'.$count0,$status)
        //         -> getStyle('U'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('U'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('U'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //         -> setCellValue('V'.$count0,$reasonrts)
        //         -> getStyle('V'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('V'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('V'.$count0)
        //         -> getFill()
        //         -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        //         -> getStartColor()
        //         -> setRGB('ffff00');

        // $excel -> setActiveSheetIndex(0) 
        //         -> setCellValue('W'.$count0,$remarks)
        //         -> getStyle('W'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('W'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('W'.$count0)
        //         -> applyFromArray($border);

        // $excel -> setActiveSheetIndex(0) 
        //         -> setCellValue('X'.$count0,$datereported)
        //         -> getStyle('X'.$count0)
        //         -> applyFromArray($topHeaderStyleFontBlackCalibri);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('X'.$count0)
        //         -> applyFromArray($centerText);

        // $excel -> setActiveSheetIndex(0) 
        //         -> getStyle('X'.$count0)
        //         -> applyFromArray($border);

    }

    $excel->getActiveSheet()->fromArray($toInsert, null, 'A2');
    
}


// SECOND SHEET

// Create a new worksheet, after the default sheet
$excel->createSheet();

//SHEET NAME

$excel -> setActiveSheetIndex(1) 
        -> setTitle('Summary Report');

//STYLE

$topHeaderStyleFontBlackBoldCalibri1 = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => '000000'),
        'size'  => 11,
        'name'  => 'Calibri'
    ));

$topHeaderStyleFontBlackCalibri1 = array(
    'font'  => array(
        'bold'  => false,
        'color' => array('rgb' => '000000'),
        'size'  => 11,
        'name'  => 'Calibri'
    ));

$tableHeaderStyleFontWhiteBoldCalibri1 = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => 'FFFFFF'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$tableHeaderStyleFontWhiteCalibri1 = array(
    'font'  => array(
        'bold'  => false,
        'color' => array('rgb' => 'FFFFFF'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$tableHeaderStyleFontRedCalibri1 = array(
    'font'  => array(
        'bold'  => false,
        'color' => array('rgb' => 'ff0000'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));


//TOP HEADER

$excel -> setActiveSheetIndex(1) 
        -> mergeCells('A1:E1');

$excel -> setActiveSheetIndex(1) 
        -> setCellValue('A1','Summary Report: '.$sender.' '.$deltype) 
        -> getStyle('A1')
        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri1);

$excel -> setActiveSheetIndex(1) 
        -> mergeCells('A2:E2');

$excel -> setActiveSheetIndex(1) 
        -> setCellValue('A2','Pick Of Month: '.$month.' '.$year) 
        -> getStyle('A2')
        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri1);

//TABLE HEADER

$excel -> setActiveSheetIndex(1) 
        -> setCellValue('A4','CYCLECODE')
        -> getStyle('A4')
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri1);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('A4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('A4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('A4')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(1) 
        -> setCellValue('B4','Date Of Pick-Up')
        -> getStyle('B4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri1);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('B4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('B4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('B4')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(1) 
        -> setCellValue('C4','Volume')
        -> getStyle('C4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri1);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('C4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('C4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('C4')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(1) 
        -> setCellValue('D4','Delivered')
        -> getStyle('D4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri1);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('D4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('D4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('D4')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(1) 
        -> setCellValue('E4','RTS')
        -> getStyle('E4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri1);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('E4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('E4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('E4')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(1) 
        -> setCellValue('F4',"Out Of Scope\n(NCR)")
        -> getStyle('F4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri1);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('F4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('F4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('F4')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('F4')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(1) 
        -> setCellValue('G4',"Out Of Scope\n(Provincial)")
        -> getStyle('G4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri1);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('G4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('G4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('G4')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('G4')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(1) 
        -> setCellValue('H4',"Delivered\n(NCR)")
        -> getStyle('H4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri1);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('H4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('H4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('H4')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('H4')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(1) 
        -> setCellValue('I4',"Delivered\n(Provincial)")
        -> getStyle('I4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri1);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('I4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('I4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('I4')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('I4')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(1) 
        -> setCellValue('J4','No Status')
        -> getStyle('J4')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri1);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('J4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('J4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('J4')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(1) 
        -> setCellValue('K4','Balance')
        -> getStyle('K4')
        -> applyFromArray($tableHeaderStyleFontRedCalibri1);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('K4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('K4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(1) 
        -> getStyle('K4')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel->setActiveSheetIndex(1)->getRowDimension('4')->setRowHeight(-1);
$excel->setActiveSheetIndex(1)->getColumnDimension('A')->setWidth(12);
$excel->setActiveSheetIndex(1)->getColumnDimension('B')->setWidth(14);
$excel->setActiveSheetIndex(1)->getColumnDimension('C')->setWidth(9);
$excel->setActiveSheetIndex(1)->getColumnDimension('D')->setWidth(11);
$excel->setActiveSheetIndex(1)->getColumnDimension('E')->setWidth(8);
$excel->setActiveSheetIndex(1)->getColumnDimension('F')->setWidth(13);
$excel->setActiveSheetIndex(1)->getColumnDimension('G')->setWidth(13);
$excel->setActiveSheetIndex(1)->getColumnDimension('H')->setWidth(13);
$excel->setActiveSheetIndex(1)->getColumnDimension('I')->setWidth(13);
$excel->setActiveSheetIndex(1)->getColumnDimension('J')->setWidth(11);
$excel->setActiveSheetIndex(1)->getColumnDimension('K')->setWidth(9);

// DATA LOOP

$countStart = 4;
$footerCount = $countSummary+4;
$totalCountNCR = 0;
$totalCountPROV = 0;
$totalVolume1 = 0;
$totalDelivered1 = 0;
$totalRTS1 = 0;
$totalOOSNCR1 = 0;
$totalOOSPROV1 = 0;
$totalDeliveredNCR1 = 0;
$totalDeliveredPROV1 = 0;
$totalBalance1 = 0;
for($i1=0; $i1<$countSummary; $i1++) {
    $summaryArray = explode(',', $_POST['checkboxsummary'][$i1]);
    $cyclecode = $summaryArray[0];
    $pud = $summaryArray[1];
    $countStart++;

    // COUNT NCR VOLUME START

    $count_total_ncr = 0;

    $sql_total_ncr = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE' OR record_area = 'MARIKINA' OR record_area = 'MUNTINLUPA' OR record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_total_ncr = mysqli_query($db_conn, $sql_total_ncr);

    if($query_total_ncr) {
        while ($row = mysqli_fetch_array($query_total_ncr, MYSQLI_ASSOC)) {
            $count_total_ncr = $row["counter"];
        }
    }

    $totalCountNCR = $totalCountNCR+$count_total_ncr;

    // COUNT NCR VOLUME END

    // COUNT PROVINCIAL VOLUME START

    $count_total_prov = 0;

    $sql_total_prov = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'RIZAL' OR record_area = 'LAGUNA') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_total_prov = mysqli_query($db_conn, $sql_total_prov);

    if($query_total_prov) {
        while ($row = mysqli_fetch_array($query_total_prov, MYSQLI_ASSOC)) {
            $count_total_prov = $row["counter"];
        }
    }

    $totalCountPROV = $totalCountPROV+$count_total_prov;

    // COUNT PROVINCIAL VOLUME END

    $excel -> setActiveSheetIndex(1) 
        -> setCellValue('A'.$countStart, $cyclecode)
        -> getStyle('A'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri1);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('A'.$countStart)
            -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('A'.$countStart)
            -> applyFromArray($border);

    // DIVIDER

    $excel -> setActiveSheetIndex(1) 
        -> setCellValue('B'.$countStart, $pud)
        -> getStyle('B'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri1);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('B'.$countStart)
            -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('B'.$countStart)
            -> applyFromArray($border);

    // DIVIDER

    $count_volume = 0;

    $sql_volume = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_volume = mysqli_query($db_conn, $sql_volume);
    if($query_volume) {
        while ($row = mysqli_fetch_array($query_volume, MYSQLI_ASSOC)) {
            $count_volume = $row["counter"];
        }
    }

    $totalVolume1 = $totalVolume1+$count_volume;

    $excel -> setActiveSheetIndex(1) 
        -> setCellValue('C'.$countStart, $count_volume)
        -> getStyle('C'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri1);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('C'.$countStart)
            -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('C'.$countStart)
            -> applyFromArray($border);

    // DIVIDER

    $count_delivered = 0;

    $sql_delivered = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_status = 'DELIVERED' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_delivered = mysqli_query($db_conn, $sql_delivered);
    if($query_delivered) {
        while ($row = mysqli_fetch_array($query_delivered, MYSQLI_ASSOC)) {
            $count_delivered = $row["counter"];
        }
    }

    $totalDelivered1 = $totalDelivered1+$count_delivered;

    $excel -> setActiveSheetIndex(1) 
        -> setCellValue('D'.$countStart,$count_delivered)
        -> getStyle('D'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri1);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('D'.$countStart)
            -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('D'.$countStart)
            -> applyFromArray($border);

    // DIVIDER

    $count_rts = 0;

    $sql_rts = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_status = 'RTS' AND record_reason_rts != 'OUT OF SCOPE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_rts = mysqli_query($db_conn, $sql_rts);
    if($query_rts) {
        while ($row = mysqli_fetch_array($query_rts, MYSQLI_ASSOC)) {
            $count_rts = $row["counter"];
        }
    }

    $totalRTS1 = $totalRTS1+$count_rts;

    $excel -> setActiveSheetIndex(1) 
        -> setCellValue('E'.$countStart,$count_rts)
        -> getStyle('E'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri1);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('E'.$countStart)
            -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('E'.$countStart)
            -> applyFromArray($border);

    // DIVIDER

    $count_oos_ncr = 0;

    $sql_oos_ncr = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_status = 'RTS' AND record_reason_rts = 'OUT OF SCOPE' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE' OR record_area = 'MARIKINA' OR record_area = 'MUNTINLUPA' OR record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_oos_ncr = mysqli_query($db_conn, $sql_oos_ncr);
    if($query_oos_ncr) {
        while ($row = mysqli_fetch_array($query_oos_ncr, MYSQLI_ASSOC)) {
            $count_oos_ncr = $row["counter"];
        }
    }

    $totalOOSNCR1 = $totalOOSNCR1+$count_oos_ncr;

    $excel -> setActiveSheetIndex(1) 
        -> setCellValue('F'.$countStart,$count_oos_ncr)
        -> getStyle('F'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri1);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('F'.$countStart)
            -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('F'.$countStart)
            -> applyFromArray($border);

    // DIVIDER

    $count_oos_prov = 0;

    $sql_oos_prov = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_status = 'RTS' AND record_reason_rts = 'OUT OF SCOPE' AND (record_area = 'LAGUNA' OR record_area = 'RIZAL') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_oos_prov = mysqli_query($db_conn, $sql_oos_prov);
    if($query_oos_prov) {
        while ($row = mysqli_fetch_array($query_oos_prov, MYSQLI_ASSOC)) {
            $count_oos_prov = $row["counter"];
        }
    }

    $totalOOSPROV1 = $totalOOSPROV1+$count_oos_prov;

    $excel -> setActiveSheetIndex(1) 
        -> setCellValue('G'.$countStart,$count_oos_prov)
        -> getStyle('G'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri1);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('G'.$countStart)
            -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('G'.$countStart)
            -> applyFromArray($border);

    // DIVIDER

    $count_delivered_ncr = 0;

    $sql_delivered_ncr = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_status = 'DELIVERED' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE' OR record_area = 'MARIKINA' OR record_area = 'MUNTINLUPA' OR record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_delivered_ncr = mysqli_query($db_conn, $sql_delivered_ncr);
    if($query_delivered_ncr) {
        while ($row = mysqli_fetch_array($query_delivered_ncr, MYSQLI_ASSOC)) {
            $count_delivered_ncr = $row["counter"];
        }
    }

    $totalDeliveredNCR1 = $totalDeliveredNCR1+$count_delivered_ncr;

    $excel -> setActiveSheetIndex(1) 
        -> setCellValue('H'.$countStart,$count_delivered_ncr)
        -> getStyle('H'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri1);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('H'.$countStart)
            -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('H'.$countStart)
            -> applyFromArray($border);

    // DIVIDER

    $count_delivered_prov = 0;

    $sql_delivered_prov = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_status = 'DELIVERED' AND (record_area = 'LAGUNA' OR record_area = 'RIZAL') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_delivered_prov = mysqli_query($db_conn, $sql_delivered_prov);
    if($query_delivered_prov) {
        while ($row = mysqli_fetch_array($query_delivered_prov, MYSQLI_ASSOC)) {
            $count_delivered_prov = $row["counter"];
        }
    }

    $totalDeliveredPROV1 = $totalDeliveredPROV1+$count_delivered_prov;

    $excel -> setActiveSheetIndex(1) 
        -> setCellValue('I'.$countStart,$count_delivered_prov)
        -> getStyle('I'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri1);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('I'.$countStart)
            -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('I'.$countStart)
            -> applyFromArray($border);

    // DIVIDER

    $excel -> setActiveSheetIndex(1) 
        -> setCellValue('J'.$countStart,'0')
        -> getStyle('J'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri1);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('J'.$countStart)
            -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('J'.$countStart)
            -> applyFromArray($border);

    // DIVIDER

    $count_balance = 0;

    $sql_balance = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_status = '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_balance = mysqli_query($db_conn, $sql_balance);
    if($query_balance) {
        while ($row = mysqli_fetch_array($query_balance, MYSQLI_ASSOC)) {
            $count_balance = $row["counter"];
        }
    }

    $totalBalance1 = $totalBalance1+$count_balance;

    $excel -> setActiveSheetIndex(1) 
            -> setCellValue('K'.$countStart, $count_balance)
            -> getStyle('K'.$countStart)
            -> applyFromArray($tableHeaderStyleFontRedCalibri1);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('K'.$countStart)
            -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('K'.$countStart)
            -> applyFromArray($border);

    $excel -> setActiveSheetIndex(1) 
            -> getStyle('K'.$countStart)
            -> getFill()
            -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            -> getStartColor()
            -> setRGB('ffff00');


    // 4 BLANK SPACE START

    if($i1 == $countSummary-1) {
        for($j=0; $j<8; $j++){
            $countStart++;

            if($j < 4) {

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('A'.$countStart, '')
                    -> getStyle('A'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('A'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('A'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('B'.$countStart, '')
                    -> getStyle('B'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('B'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('B'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('C'.$countStart, '')
                    -> getStyle('C'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackBoldCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('C'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('C'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('D'.$countStart, '')
                    -> getStyle('D'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('D'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('D'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('E'.$countStart, '')
                    -> getStyle('E'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('E'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('E'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('F'.$countStart, '')
                    -> getStyle('F'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('F'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('F'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('G'.$countStart, '')
                    -> getStyle('G'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('G'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('G'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('H'.$countStart, '')
                    -> getStyle('H'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('H'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('H'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('I'.$countStart, '')
                    -> getStyle('I'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('I'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('I'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('J'.$countStart, '')
                    -> getStyle('J'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('J'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('J'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                        -> setCellValue('K'.$countStart,'')
                        -> getStyle('K'.$countStart)
                        -> applyFromArray($tableHeaderStyleFontRedCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('K'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('K'.$countStart)
                        -> applyFromArray($border);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('K'.$countStart)
                        -> getFill()
                        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                        -> getStartColor()
                        -> setRGB('ffff00');

            } else if ($j == 4) {

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('A'.$countStart, 'NCR')
                    -> getStyle('A'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackBoldCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('A'.$countStart)
                        -> applyFromArray($border);

                 $excel -> setActiveSheetIndex(1) 
                        -> getStyle('A'.$countStart)
                        -> getFill()
                        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                        -> getStartColor()
                        -> setRGB('d9d9d9');

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('B'.$countStart, '')
                    -> getStyle('B'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('B'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('B'.$countStart)
                        -> applyFromArray($border);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('B'.$countStart)
                        -> getFill()
                        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                        -> getStartColor()
                        -> setRGB('d9d9d9');

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('C'.$countStart, $totalCountNCR)
                    -> getStyle('C'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackBoldCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('C'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('C'.$countStart)
                        -> applyFromArray($border);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('C'.$countStart)
                        -> getFill()
                        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                        -> getStartColor()
                        -> setRGB('d9d9d9');

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('D'.$countStart, '')
                    -> getStyle('D'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('D'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('D'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('E'.$countStart, '')
                    -> getStyle('E'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('E'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('E'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('F'.$countStart, '')
                    -> getStyle('F'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('F'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('F'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('G'.$countStart, '')
                    -> getStyle('G'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('G'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('G'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('H'.$countStart, '')
                    -> getStyle('H'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('H'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('H'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('I'.$countStart, '')
                    -> getStyle('I'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('I'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('I'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('J'.$countStart, '')
                    -> getStyle('J'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('J'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('J'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                        -> setCellValue('K'.$countStart,'')
                        -> getStyle('K'.$countStart)
                        -> applyFromArray($tableHeaderStyleFontRedCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('K'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('K'.$countStart)
                        -> applyFromArray($border);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('K'.$countStart)
                        -> getFill()
                        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                        -> getStartColor()
                        -> setRGB('ffff00');

            } else if ($j == 5) {

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('A'.$countStart, 'PROVINCIAL')
                    -> getStyle('A'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackBoldCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('A'.$countStart)
                        -> applyFromArray($border);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('A'.$countStart)
                        -> getFill()
                        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                        -> getStartColor()
                        -> setRGB('d9d9d9');

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('B'.$countStart, '')
                    -> getStyle('B'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('B'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('B'.$countStart)
                        -> applyFromArray($border);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('B'.$countStart)
                        -> getFill()
                        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                        -> getStartColor()
                        -> setRGB('d9d9d9');

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('C'.$countStart, $totalCountPROV)
                    -> getStyle('C'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackBoldCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('C'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('C'.$countStart)
                        -> applyFromArray($border);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('C'.$countStart)
                        -> getFill()
                        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                        -> getStartColor()
                        -> setRGB('d9d9d9');

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('D'.$countStart, '')
                    -> getStyle('D'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('D'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('D'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('E'.$countStart, '')
                    -> getStyle('E'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('E'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('E'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('F'.$countStart, '')
                    -> getStyle('F'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('F'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('F'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('G'.$countStart, '')
                    -> getStyle('G'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('G'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('G'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('H'.$countStart, '')
                    -> getStyle('H'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('H'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('H'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('I'.$countStart, '')
                    -> getStyle('I'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('I'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('I'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('J'.$countStart, '')
                    -> getStyle('J'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('J'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('J'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                        -> setCellValue('K'.$countStart,'')
                        -> getStyle('K'.$countStart)
                        -> applyFromArray($tableHeaderStyleFontRedCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('K'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('K'.$countStart)
                        -> applyFromArray($border);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('K'.$countStart)
                        -> getFill()
                        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                        -> getStartColor()
                        -> setRGB('ffff00');
                
            } else if ($j == 6) {

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('A'.$countStart, 'Total')
                    -> getStyle('A'.$countStart)
                    -> applyFromArray($tableHeaderStyleFontRedCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('A'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('A'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('B'.$countStart, '')
                    -> getStyle('B'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('B'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('B'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('C'.$countStart, $totalCountNCR+$totalCountPROV)
                    -> getStyle('C'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackBoldCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('C'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('C'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('D'.$countStart, $totalDelivered1)
                    -> getStyle('D'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('D'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('D'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('E'.$countStart, $totalRTS1)
                    -> getStyle('E'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('E'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('E'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('F'.$countStart, $totalOOSNCR1)
                    -> getStyle('F'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('F'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('F'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('G'.$countStart, $totalOOSPROV1)
                    -> getStyle('G'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('G'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('G'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('H'.$countStart, $totalDeliveredNCR1)
                    -> getStyle('H'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('H'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('H'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('I'.$countStart, $totalDeliveredPROV1)
                    -> getStyle('I'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('I'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('I'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('J'.$countStart, '0')
                    -> getStyle('J'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('J'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('J'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                        -> setCellValue('K'.$countStart, $totalBalance1)
                        -> getStyle('K'.$countStart)
                        -> applyFromArray($tableHeaderStyleFontRedCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('K'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('K'.$countStart)
                        -> applyFromArray($border);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('K'.$countStart)
                        -> getFill()
                        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                        -> getStartColor()
                        -> setRGB('ffff00');
                
            } else if ($j == 7) {

                // PERCENTAGE START

                $totalCount = $totalCountNCR+$totalCountPROV;

                $delivered = 0.00;
                $rts = 0.00;
                $oosncr = 0.00;
                $oosprov = 0.00;
                $deliveredncr = 0.00;
                $deliveredprov = 0.00;

                if($totalDelivered1 > 0) {
                    $delivered = ($totalDelivered1/$totalCount)*100;
                    $delivered = number_format($delivered, 2);
                }

                if($totalRTS1 > 0) {
                    $rts = ($totalRTS1/$totalCount)*100;
                    $rts = number_format($rts, 2);
                }

                if($totalOOSNCR1 > 0) {
                    $oosncr = ($totalOOSNCR1/$totalCount)*100;
                    $oosncr = number_format($oosncr, 2);
                }

                if($totalOOSPROV1 > 0) {
                    $oosprov = ($totalOOSPROV1/$totalCount)*100;
                    $oosprov = number_format($oosprov, 2);
                }

                if($totalDeliveredNCR1 > 0) {
                    $deliveredncr = ($totalDeliveredNCR1/$totalCount)*100;
                    $deliveredncr = number_format($deliveredncr, 2);
                }

                if($totalDeliveredPROV1 > 0) {
                    $deliveredprov = ($totalDeliveredPROV1/$totalCount)*100;
                    $deliveredprov = number_format($deliveredprov, 2);
                }
                // PERCENTAGE END


                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('A'.$countStart, 'Percentage')
                    -> getStyle('A'.$countStart)
                    -> applyFromArray($tableHeaderStyleFontRedCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('A'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('A'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('B'.$countStart, '')
                    -> getStyle('B'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('B'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('B'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('C'.$countStart, '')
                    -> getStyle('C'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackBoldCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('C'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('C'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('D'.$countStart, $delivered.'%')
                    -> getStyle('D'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('D'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('D'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('E'.$countStart, $rts.'%')
                    -> getStyle('E'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('E'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('E'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('F'.$countStart, $oosncr.'%')
                    -> getStyle('F'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('F'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('F'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('G'.$countStart, $oosprov.'%')
                    -> getStyle('G'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('G'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('G'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('H'.$countStart, $deliveredncr.'%')
                    -> getStyle('H'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('H'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('H'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('I'.$countStart, $deliveredprov.'%')
                    -> getStyle('I'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('I'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('I'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                    -> setCellValue('J'.$countStart, '0.00%')
                    -> getStyle('J'.$countStart)
                    -> applyFromArray($topHeaderStyleFontBlackCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('J'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('J'.$countStart)
                        -> applyFromArray($border);

                // DIVIDER

                $excel -> setActiveSheetIndex(1) 
                        -> setCellValue('K'.$countStart,'')
                        -> getStyle('K'.$countStart)
                        -> applyFromArray($tableHeaderStyleFontRedCalibri1);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('K'.$countStart)
                        -> applyFromArray($centerText);

                $excel -> setActiveSheetIndex(1) 
                        -> getStyle('K'.$countStart)
                        -> applyFromArray($border);

                $excel -> setActiveSheetIndex(1) 
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

// THIRD SHEET

// Create a new worksheet, after the default sheet
$excel->createSheet();

//SHEET NAME

$excel -> setActiveSheetIndex(2) 
        -> setTitle('Days of Delivery');

//STYLE

$topHeaderStyleFontBlackBoldCalibri2 = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => '000000'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$topHeaderStyleFontBlackCalibri2 = array(
    'font'  => array(
        'bold'  => false,
        'color' => array('rgb' => '000000'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$tableHeaderStyleFontWhiteBoldCalibri2 = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => 'FFFFFF'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$tableHeaderStyleFontWhiteCalibri2 = array(
    'font'  => array(
        'bold'  => false,
        'color' => array('rgb' => 'FFFFFF'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$tableHeaderStyleFontBlackBoldCalibri2 = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => '000000'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$tableHeaderStyleFontBlackCalibri2 = array(
    'font'  => array(
        'bold'  => false,
        'color' => array('rgb' => '000000'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$tableHeaderStyleFontRedBoldCalibri2 = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => 'ff0000'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$tableHeaderStyleFontRedCalibri2 = array(
    'font'  => array(
        'bold'  => false,
        'color' => array('rgb' => 'ff0000'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$tableHeaderStyleFontYellowBoldCalibri2 = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => 'ffff00'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$tableHeaderStyleFontYellowBoldCalibri2 = array(
    'font'  => array(
        'bold'  => false,
        'color' => array('rgb' => 'ffff00'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));


//TOP HEADER

$totalVolume = 0;
$totalDay1 = 0;
$totalDay2 = 0;
$totalDay3 = 0;
$totalDay4 = 0;
$totalDay5 = 0;
$totalDay6 = 0;
$totalDay7 = 0;
$totalDay8 = 0;
$totalDay9 = 0;
$totalDay10 = 0;
$totalDay11 = 0;
$totalDay12 = 0;
$totalDay13 = 0;
$totalDay14 = 0;
$totalDay15 = 0;
$totalDelivered = 0;
$totalRTS = 0;
$totalOOS = 0;
$totalBalance = 0;
$totalVolumeInvoice = 0;
$totalAmount = 0;
$ncrTotalDay1 = 0;
$ncrTotalDay2 = 0;
$ncrTotalDay3 = 0;
$ncrTotalDay4 = 0;
$ncrTotalDay5 = 0;
$ncrTotalDay6 = 0;
$ncrTotalDay7 = 0;
$ncrTotalDay8 = 0;
$ncrTotalDay9 = 0;
$ncrTotalDay10 = 0;
$ncrTotalDay11 = 0;
$ncrTotalDay12 = 0;
$ncrTotalDay13 = 0;
$ncrTotalDay14 = 0;
$ncrTotalDay15 = 0;
$provTotalDay1 = 0;
$provTotalDay2 = 0;
$provTotalDay3 = 0;
$provTotalDay4 = 0;
$provTotalDay5 = 0;
$provTotalDay6 = 0;
$provTotalDay7 = 0;
$provTotalDay8 = 0;
$provTotalDay9 = 0;
$provTotalDay10 = 0;
$provTotalDay11 = 0;
$provTotalDay12 = 0;
$provTotalDay13 = 0;
$provTotalDay14 = 0;
$provTotalDay15 = 0;

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('A1:Z1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('A1',strtoupper($month).' '.strtoupper($sender).'('.strtoupper($deltype).') PERFORMANCE EVALUATION') 
        -> getStyle('A1')
        -> applyFromArray($topHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2)
        -> getStyle('A1')->getFont()->setSize(24);

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('A3:Z3');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('A3','F L O W  OF  D E L I V E R Y') 
        -> getStyle('A3')
        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A3')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2)
        -> getStyle('A3')->getFont()->setSize(16);

//TABLE HEADER

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('A4:A8');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('A4','SENDER')
        -> getStyle('A4:A8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A4:A8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A4:A8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A4:A8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('B4:B8');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('B4','DELIVERY TYPE')
        -> getStyle('B4:B8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('B4:B8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('B4:B8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('B4:B8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('C4:C8');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('C4','AREA')
        -> getStyle('C4:C8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('C4:C8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('C4:C8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('C4:C8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('D4:D8');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('D4','MONTH')
        -> getStyle('D4:D8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('D4:D8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('D4:D8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('D4:D8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('E4:E8');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('E4','CYCLE CODE')
        -> getStyle('E4:E8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('E4:E8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('E4:E8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('E4:E8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('F4:F8');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('F4','PICK-UP DATE')
        -> getStyle('F4:F8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('F4:F8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('F4:F8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('F4:F8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('F4:F8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('G4:G8');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('G4','VOLUME')
        -> getStyle('G4:G8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G4:G8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G4:G8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G4:G8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G4:G8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('H4:V4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('H4','DELIVERY TIMELINE')
        -> getStyle('H4:V4')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H4:V4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H4:V4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H4:V4')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H4:V4')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('H5:Q5');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('H5','ON TIME DELIVERY FOR PROVINCIAL')
        -> getStyle('H5:Q5')
        -> applyFromArray($tableHeaderStyleFontYellowBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H5:Q5')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H5:Q5')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H5:Q5')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H5:Q5')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('0070c0');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('R5:V5');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('R5','LATE DELIVERY FOR PROVINCIAL')
        -> getStyle('R5:V5')
        -> applyFromArray($tableHeaderStyleFontYellowBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R5:V5')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R5:V5')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R5:V5')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R5:V5')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('0070c0');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('H6','DAY 1')
        -> getStyle('H6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('I6','DAY 2')
        -> getStyle('I6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('J6','DAY 3')
        -> getStyle('J6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('K6','DAY 4')
        -> getStyle('K6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('L6','DAY 5')
        -> getStyle('L6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('M6','DAY 6')
        -> getStyle('M6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('N6','DAY 7')
        -> getStyle('N6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('O6','DAY 8')
        -> getStyle('O6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('P6','DAY 9')
        -> getStyle('P6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Q6','DAY 10')
        -> getStyle('Q6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('R6','DAY 11')
        -> getStyle('R6')
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('S6','DAY 12')
        -> getStyle('S6')
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('T6','DAY 13')
        -> getStyle('T6')
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('U6','DAY 14')
        -> getStyle('U6')
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('V6','DAY 15')
        -> getStyle('V6')
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('H7:L7');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('H7','ON TIME DELIVERY FOR NCR')
        -> getStyle('H7:L7')
        -> applyFromArray($tableHeaderStyleFontYellowBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H7:L7')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H7:L7')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H7:L7')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H7:L7')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('0070c0');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('M7:V7');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('M7','LATE DELIVERY FOR NCR')
        -> getStyle('M7:V7')
        -> applyFromArray($tableHeaderStyleFontYellowBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M7:V7')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M7:V7')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M7:V7')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M7:V7')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('0070c0');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('H8','DAY 1')
        -> getStyle('H8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('I8','DAY 2')
        -> getStyle('I8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('J8','DAY 3')
        -> getStyle('J8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('K8','DAY 4')
        -> getStyle('K8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('L8','DAY 5')
        -> getStyle('L8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('M8','DAY 6')
        -> getStyle('M8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('N8','DAY 7')
        -> getStyle('N8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('O8','DAY 8')
        -> getStyle('O8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('P8','DAY 9')
        -> getStyle('P8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Q8','DAY 10')
        -> getStyle('Q8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('R8','DAY 11')
        -> getStyle('R8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('S8','DAY 12')
        -> getStyle('S8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('T8','DAY 13')
        -> getStyle('T8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('U8','DAY 14')
        -> getStyle('U8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('V8','DAY 15')
        -> getStyle('V8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('W4:W8');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('W4',"DELIVERED")
        -> getStyle('W4:W8')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W4:W8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W4:W8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W4:W8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W4:W8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('X4:X8');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('X4','RTS')
        -> getStyle('X4:X8')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X4:X8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X4:X8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X4:X8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('Y4:Y8');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Y4','OUT OF SCOPE')
        -> getStyle('Y4:Y8')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y4:Y8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y4:Y8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y4:Y8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('Z4:Z8');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Z4','BALANCE')
        -> getStyle('Z4:Z8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z4:Z8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z4:Z8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z4:Z8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z4:Z8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('AA4:AA8');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AA4',"REPORT STATUS")
        -> getStyle('AA4:AA8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AA4:AA8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AA4:AA8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AA4:AA8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('AB4:AB8');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AB4','AREA')
        -> getStyle('AB4:AB8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AB4:AB8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AB4:AB8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AB4:AB8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('AC4:AC8');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AC4',"TOTAL VOLUME INVOICE")
        -> getStyle('AC4:AC8')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC4:AC8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC4:AC8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC4:AC8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC4:AC8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('AD4:AD8');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AD4','PRICE')
        -> getStyle('AD4:AD8')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD4:AD8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD4:AD8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD4:AD8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('AE4:AE8');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AE4','AMOUNT')
        -> getStyle('AE4:AE8')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE4:AE8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE4:AE8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE4:AE8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');


$excel -> setActiveSheetIndex(2) 
        -> mergeCells('A9:AE9');

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A9:AE9')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel->setActiveSheetIndex(2)->getRowDimension('4')->setRowHeight(-1);
$excel->setActiveSheetIndex(2)->getRowDimension('5')->setRowHeight(-1);
$excel->setActiveSheetIndex(2)->getRowDimension('6')->setRowHeight(-1);
$excel->setActiveSheetIndex(2)->getRowDimension('7')->setRowHeight(-1);
$excel->setActiveSheetIndex(2)->getRowDimension('8')->setRowHeight(-1);
$excel->setActiveSheetIndex(2)->getColumnDimension('A')->setWidth(9);
$excel->setActiveSheetIndex(2)->getColumnDimension('B')->setWidth(15);
$excel->setActiveSheetIndex(2)->getColumnDimension('C')->setWidth(8);
$excel->setActiveSheetIndex(2)->getColumnDimension('D')->setWidth(10);
$excel->setActiveSheetIndex(2)->getColumnDimension('E')->setWidth(13);
$excel->setActiveSheetIndex(2)->getColumnDimension('F')->setWidth(13);
$excel->setActiveSheetIndex(2)->getColumnDimension('G')->setWidth(9);
$excel->setActiveSheetIndex(2)->getColumnDimension('W')->setWidth(11);
$excel->setActiveSheetIndex(2)->getColumnDimension('X')->setWidth(9);
$excel->setActiveSheetIndex(2)->getColumnDimension('Y')->setWidth(13);
$excel->setActiveSheetIndex(2)->getColumnDimension('Z')->setWidth(11);
$excel->setActiveSheetIndex(2)->getColumnDimension('AA')->setWidth(14);
$excel->setActiveSheetIndex(2)->getColumnDimension('AB')->setWidth(12);
$excel->setActiveSheetIndex(2)->getColumnDimension('AC')->setWidth(14);
$excel->setActiveSheetIndex(2)->getColumnDimension('AD')->setWidth(12);
$excel->setActiveSheetIndex(2)->getColumnDimension('AE')->setWidth(13);

// DATA LOOP

$countSummary = count($_POST['checkboxsummary']);
$count = 10;

for($i=0; $i<$countSummary; $i++) {
    $summaryArray = explode(',', $_POST['checkboxsummary'][$i]);
    $cyclecode = $summaryArray[0];
    $pud = $summaryArray[1];

    $reportStatusPQUE = "Incomplete";
    $reportStatusLPNAS = "Incomplete";
    $reportStatusMLUPA = "Incomplete";
    $reportStatusMKNA = "Incomplete";
    $reportStatusRIZAL = "Incomplete";
    $reportStatusLAGUNA = "Incomplete";

    $mergeNumber = $count+5;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('A'.$count,strtoupper($sender))
        -> getStyle('A'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('B'.$count,strtoupper($deltype))
        -> getStyle('B'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('B'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('B'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('C'.$count,'PQUE')
        -> getStyle('C'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('C'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('C'.$count)
        -> applyFromArray($border);

//

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('D'.$count.':D'.$mergeNumber);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('D'.$count,strtoupper($month))
        -> getStyle('D'.$count.':D'.$mergeNumber)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('D'.$count.':D'.$mergeNumber)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('D'.$count.':D'.$mergeNumber)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('E'.$count.':E'.$mergeNumber);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('E'.$count,$cyclecode)
        -> getStyle('E'.$count.':E'.$mergeNumber)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('E'.$count.':E'.$mergeNumber)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('E'.$count.':E'.$mergeNumber)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('F'.$count.':F'.$mergeNumber);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('F'.$count,$pud)
        -> getStyle('F'.$count.':F'.$mergeNumber)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('F'.$count.':F'.$mergeNumber)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('F'.$count.':F'.$mergeNumber)
        -> applyFromArray($border);

//

$count_total_pque = 0;

$sql_total_pque = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_pque = mysqli_query($db_conn, $sql_total_pque);
if($query_total_pque) {
    while ($row = mysqli_fetch_array($query_total_pque, MYSQLI_ASSOC)) {
            $count_total_pque = $row["counter"];
        }
}

$totalVolume = $totalVolume+$count_total_pque;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('G'.$count,$count_total_pque)
        -> getStyle('G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G'.$count)
        -> applyFromArray($border);

//

$timestamp = strtotime($pud);
$puddatetime = date("Y-m-d H:i:s", $timestamp);

$count_total_pque_1day = 0;
$count_total_pque_2day = 0;
$count_total_pque_3day = 0;
$count_total_pque_4day = 0;
$count_total_pque_5day = 0;
$count_total_pque_6day = 0;
$count_total_pque_7day = 0;
$count_total_pque_8day = 0;
$count_total_pque_9day = 0;
$count_total_pque_10day = 0;
$count_total_pque_11day = 0;
$count_total_pque_12day = 0;
$count_total_pque_13day = 0;
$count_total_pque_14day = 0;
$count_total_pque_15day = 0;

$sql_total_pque_1day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -1 DAY) AND record_status_status='1'";
$query_total_pque_1day = mysqli_query($db_conn, $sql_total_pque_1day);
if($query_total_pque_1day) {
    while ($row = mysqli_fetch_array($query_total_pque_1day, MYSQLI_ASSOC)) {
            $count_total_pque_1day = $row["counter"];
        }
}

$sql_total_pque_2day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -2 DAY) AND record_status_status='1'";
$query_total_pque_2day = mysqli_query($db_conn, $sql_total_pque_2day);
if($query_total_pque_2day) {
    while ($row = mysqli_fetch_array($query_total_pque_2day, MYSQLI_ASSOC)) {
            $count_total_pque_2day = $row["counter"];
        }
}

$sql_total_pque_3day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -3 DAY) AND record_status_status='1'";
$query_total_pque_3day = mysqli_query($db_conn, $sql_total_pque_3day);
if($query_total_pque_3day) {
    while ($row = mysqli_fetch_array($query_total_pque_3day, MYSQLI_ASSOC)) {
            $count_total_pque_3day = $row["counter"];
        }
}

$sql_total_pque_4day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -4 DAY) AND record_status_status='1'";
$query_total_pque_4day = mysqli_query($db_conn, $sql_total_pque_4day);
if($query_total_pque_4day) {
    while ($row = mysqli_fetch_array($query_total_pque_4day, MYSQLI_ASSOC)) {
            $count_total_pque_4day = $row["counter"];
        }
}

$sql_total_pque_5day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -5 DAY) AND record_status_status='1'";
$query_total_pque_5day = mysqli_query($db_conn, $sql_total_pque_5day);
if($query_total_pque_5day) {
    while ($row = mysqli_fetch_array($query_total_pque_5day, MYSQLI_ASSOC)) {
            $count_total_pque_5day = $row["counter"];
        }
}

$sql_total_pque_6day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -6 DAY) AND record_status_status='1'";
$query_total_pque_6day = mysqli_query($db_conn, $sql_total_pque_6day);
if($query_total_pque_6day) {
    while ($row = mysqli_fetch_array($query_total_pque_6day, MYSQLI_ASSOC)) {
            $count_total_pque_6day = $row["counter"];
        }
}

$sql_total_pque_7day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -7 DAY) AND record_status_status='1'";
$query_total_pque_7day = mysqli_query($db_conn, $sql_total_pque_7day);
if($query_total_pque_7day) {
    while ($row = mysqli_fetch_array($query_total_pque_7day, MYSQLI_ASSOC)) {
            $count_total_pque_7day = $row["counter"];
        }
}

$sql_total_pque_8day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -8 DAY) AND record_status_status='1'";
$query_total_pque_8day = mysqli_query($db_conn, $sql_total_pque_8day);
if($query_total_pque_8day) {
    while ($row = mysqli_fetch_array($query_total_pque_8day, MYSQLI_ASSOC)) {
            $count_total_pque_8day = $row["counter"];
        }
    }

$sql_total_pque_9day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -9 DAY) AND record_status_status='1'";
$query_total_pque_9day = mysqli_query($db_conn, $sql_total_pque_9day);
if($query_total_pque_9day) {
    while ($row = mysqli_fetch_array($query_total_pque_9day, MYSQLI_ASSOC)) {
            $count_total_pque_9day = $row["counter"];
        }
    }

$sql_total_pque_10day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -10 DAY) AND record_status_status='1'";
$query_total_pque_10day = mysqli_query($db_conn, $sql_total_pque_10day);
if($query_total_pque_10day) {
    while ($row = mysqli_fetch_array($query_total_pque_10day, MYSQLI_ASSOC)) {
            $count_total_pque_10day = $row["counter"];
        }
    }

$sql_total_pque_11day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -11 DAY) AND record_status_status='1'";
$query_total_pque_11day = mysqli_query($db_conn, $sql_total_pque_11day);
if($query_total_pque_11day) {
    while ($row = mysqli_fetch_array($query_total_pque_11day, MYSQLI_ASSOC)) {
            $count_total_pque_11day = $row["counter"];
        }
    }

$sql_total_pque_12day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -12 DAY) AND record_status_status='1'";
$query_total_pque_12day = mysqli_query($db_conn, $sql_total_pque_12day);
if($query_total_pque_12day) {
    while ($row = mysqli_fetch_array($query_total_pque_12day, MYSQLI_ASSOC)) {
            $count_total_pque_12day = $row["counter"];
        }
    }

$sql_total_pque_13day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -13 DAY) AND record_status_status='1'";
$query_total_pque_13day = mysqli_query($db_conn, $sql_total_pque_13day);
if($query_total_pque_13day) {
    while ($row = mysqli_fetch_array($query_total_pque_13day, MYSQLI_ASSOC)) {
            $count_total_pque_13day = $row["counter"];
        }
    }

$sql_total_pque_14day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -14 DAY) AND record_status_status='1'";
$query_total_pque_14day = mysqli_query($db_conn, $sql_total_pque_14day);
if($query_total_pque_14day) {
    while ($row = mysqli_fetch_array($query_total_pque_14day, MYSQLI_ASSOC)) {
            $count_total_pque_14day = $row["counter"];
        }
    }

$sql_total_pque_15day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -15 DAY) AND record_status_status='1'";
$query_total_pque_15day = mysqli_query($db_conn, $sql_total_pque_15day);
if($query_total_pque_15day) {
    while ($row = mysqli_fetch_array($query_total_pque_15day, MYSQLI_ASSOC)) {
            $count_total_pque_15day = $row["counter"];
        }
    }

$totalDay1 = $totalDay1+$count_total_pque_1day;
$totalDay2 = $totalDay2+$count_total_pque_2day;
$totalDay3 = $totalDay3+$count_total_pque_3day;
$totalDay4 = $totalDay4+$count_total_pque_4day;
$totalDay5 = $totalDay5+$count_total_pque_5day;
$totalDay6 = $totalDay6+$count_total_pque_6day;
$totalDay7 = $totalDay7+$count_total_pque_7day;
$totalDay8 = $totalDay8+$count_total_pque_8day;
$totalDay9 = $totalDay9+$count_total_pque_9day;
$totalDay10 = $totalDay10+$count_total_pque_10day;
$totalDay11 = $totalDay11+$count_total_pque_11day;
$totalDay12 = $totalDay12+$count_total_pque_12day;
$totalDay13 = $totalDay13+$count_total_pque_13day;
$totalDay14 = $totalDay14+$count_total_pque_14day;
$totalDay15 = $totalDay15+$count_total_pque_15day;

$ncrTotalDay1 = $ncrTotalDay1+$count_total_pque_1day;
$ncrTotalDay2 = $ncrTotalDay2+$count_total_pque_2day;
$ncrTotalDay3 = $ncrTotalDay3+$count_total_pque_3day;
$ncrTotalDay4 = $ncrTotalDay4+$count_total_pque_4day;
$ncrTotalDay5 = $ncrTotalDay5+$count_total_pque_5day;
$ncrTotalDay6 = $ncrTotalDay6+$count_total_pque_6day;
$ncrTotalDay7 = $ncrTotalDay7+$count_total_pque_7day;
$ncrTotalDay8 = $ncrTotalDay8+$count_total_pque_8day;
$ncrTotalDay9 = $ncrTotalDay9+$count_total_pque_9day;
$ncrTotalDay10 = $ncrTotalDay10+$count_total_pque_10day;
$ncrTotalDay11 = $ncrTotalDay11+$count_total_pque_11day;
$ncrTotalDay12 = $ncrTotalDay12+$count_total_pque_12day;
$ncrTotalDay13 = $ncrTotalDay13+$count_total_pque_13day;
$ncrTotalDay14 = $ncrTotalDay14+$count_total_pque_14day;
$ncrTotalDay15 = $ncrTotalDay15+$count_total_pque_15day;


$excel -> setActiveSheetIndex(2) 
        -> setCellValue('H'.$count,$count_total_pque_1day)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('I'.$count,$count_total_pque_2day)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('J'.$count,$count_total_pque_3day)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('K'.$count,$count_total_pque_4day)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('L'.$count,$count_total_pque_5day)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('M'.$count,$count_total_pque_6day)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('N'.$count,$count_total_pque_7day)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('O'.$count,$count_total_pque_8day)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('P'.$count,$count_total_pque_9day)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Q'.$count,$count_total_pque_10day)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('R'.$count,$count_total_pque_11day)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('S'.$count,$count_total_pque_12day)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('T'.$count,$count_total_pque_13day)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('U'.$count,$count_total_pque_14day)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('V'.$count,$count_total_pque_15day)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$count_total_pque_delivered = 0;

$sql_total_pque_delivered = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_status = 'DELIVERED' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_pque_delivered = mysqli_query($db_conn, $sql_total_pque_delivered);
if($query_total_pque_delivered) {
    while ($row = mysqli_fetch_array($query_total_pque_delivered, MYSQLI_ASSOC)) {
            $count_total_pque_delivered = $row["counter"];
        }
    }

$totalDelivered = $totalDelivered+$count_total_pque_delivered;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('W'.$count,$count_total_pque_delivered)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$count_total_pque_rts = 0;

$sql_total_pque_rts = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_status = 'RTS' AND record_reason_rts != 'OUT OF SCOPE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_pque_rts = mysqli_query($db_conn, $sql_total_pque_rts);
if($query_total_pque_rts) {
    while ($row = mysqli_fetch_array($query_total_pque_rts, MYSQLI_ASSOC)) {
            $count_total_pque_rts = $row["counter"];
        }
    }

$totalRTS = $totalRTS+$count_total_pque_rts;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('X'.$count,$count_total_pque_rts)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$count_total_pque_oos = 0;

$sql_total_pque_oos = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_reason_rts = 'OUT OF SCOPE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_pque_oos = mysqli_query($db_conn, $sql_total_pque_oos);
if($query_total_pque_oos) {
    while ($row = mysqli_fetch_array($query_total_pque_oos, MYSQLI_ASSOC)) {
            $count_total_pque_oos = $row["counter"];
        }
    }

$totalOOS = $totalOOS+$count_total_pque_oos;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Y'.$count,$count_total_pque_oos)
        -> getStyle('Y'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y'.$count)
        -> applyFromArray($border);

$count_total_pque_balance = 0;

$sql_total_pque_balance = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_status = '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_pque_balance = mysqli_query($db_conn, $sql_total_pque_balance);
if($query_total_pque_balance) {
    while ($row = mysqli_fetch_array($query_total_pque_balance, MYSQLI_ASSOC)) {
            $count_total_pque_balance = $row["counter"];
        }
    }

$totalBalance = $totalBalance+$count_total_pque_balance;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Z'.$count,$count_total_pque_balance)
        -> getStyle('Z'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

if($count_total_pque_balance == 0) {
    $reportStatusPQUE = "Complete";
}

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AA'.$count,$reportStatusPQUE)
        -> getStyle('AA'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AA'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AA'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AB'.$count,'PQUE')
        -> getStyle('AB'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AB'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AB'.$count)
        -> applyFromArray($border);

$count_total_pque1 = 0;

$sql_total_pque1 = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE') AND record_status != '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_pque1 = mysqli_query($db_conn, $sql_total_pque1);
if($query_total_pque1) {
    while ($row = mysqli_fetch_array($query_total_pque1, MYSQLI_ASSOC)) {
            $count_total_pque1 = $row["counter"];
        }
    }

$count_total_pque1 = $count_total_pque1-$count_total_pque_oos;

$totalVolumeInvoice = $totalVolumeInvoice+$count_total_pque1;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AC'.$count,$count_total_pque1)
        -> getStyle('AC'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AD'.$count,$pqueprice)
        -> getStyle('AD'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$amountPQUE = $pqueprice*$count_total_pque1;

$totalAmount = $totalAmount+$amountPQUE;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AE'.$count,'Php'.number_format($amountPQUE))
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);

//2ND LINE

$count = $count+1;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('A'.$count,strtoupper($sender))
        -> getStyle('A'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('B'.$count,strtoupper($deltype))
        -> getStyle('B'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('B'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('B'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('C'.$count,'LPNAS')
        -> getStyle('C'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('C'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('C'.$count)
        -> applyFromArray($border);

$count_total_lpnas = 0;

$sql_total_lpnas = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_lpnas = mysqli_query($db_conn, $sql_total_lpnas);
if($query_total_lpnas) {
    while ($row = mysqli_fetch_array($query_total_lpnas, MYSQLI_ASSOC)) {
            $count_total_lpnas = $row["counter"];
        }
    }

$totalVolume = $totalVolume+$count_total_lpnas;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('G'.$count,$count_total_lpnas)
        -> getStyle('G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G'.$count)
        -> applyFromArray($border);

$count_total_lpnas_1day = 0;
$count_total_lpnas_2day = 0;
$count_total_lpnas_3day = 0;
$count_total_lpnas_4day = 0;
$count_total_lpnas_5day = 0;
$count_total_lpnas_6day = 0;
$count_total_lpnas_7day = 0;
$count_total_lpnas_8day = 0;
$count_total_lpnas_9day = 0;
$count_total_lpnas_10day = 0;
$count_total_lpnas_11day = 0;
$count_total_lpnas_12day = 0;
$count_total_lpnas_13day = 0;
$count_total_lpnas_14day = 0;
$count_total_lpnas_15day = 0;

$sql_total_lpnas_1day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -1 DAY) AND record_status_status='1'";
$query_total_lpnas_1day = mysqli_query($db_conn, $sql_total_lpnas_1day);
if($query_total_lpnas_1day) {
    while ($row = mysqli_fetch_array($query_total_lpnas_1day, MYSQLI_ASSOC)) {
            $count_total_lpnas_1day = $row["counter"];
        }
    }

$sql_total_lpnas_2day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -2 DAY) AND record_status_status='1'";
$query_total_lpnas_2day = mysqli_query($db_conn, $sql_total_lpnas_2day);
if($query_total_lpnas_2day) {
    while ($row = mysqli_fetch_array($query_total_lpnas_2day, MYSQLI_ASSOC)) {
            $count_total_lpnas_2day = $row["counter"];
        }
    }

$sql_total_lpnas_3day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -3 DAY) AND record_status_status='1'";
$query_total_lpnas_3day = mysqli_query($db_conn, $sql_total_lpnas_3day);
if($query_total_lpnas_3day) {
    while ($row = mysqli_fetch_array($query_total_lpnas_3day, MYSQLI_ASSOC)) {
            $count_total_lpnas_3day = $row["counter"];
        }
    }

$sql_total_lpnas_4day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -4 DAY) AND record_status_status='1'";
$query_total_lpnas_4day = mysqli_query($db_conn, $sql_total_lpnas_4day);
if($query_total_lpnas_4day) {
    while ($row = mysqli_fetch_array($query_total_lpnas_4day, MYSQLI_ASSOC)) {
            $count_total_lpnas_4day = $row["counter"];
        }
    }

$sql_total_lpnas_5day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -5 DAY) AND record_status_status='1'";
$query_total_lpnas_5day = mysqli_query($db_conn, $sql_total_lpnas_5day);
if($query_total_lpnas_5day) {
    while ($row = mysqli_fetch_array($query_total_lpnas_5day, MYSQLI_ASSOC)) {
            $count_total_lpnas_5day = $row["counter"];
        }
    }

$sql_total_lpnas_6day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -6 DAY) AND record_status_status='1'";
$query_total_lpnas_6day = mysqli_query($db_conn, $sql_total_lpnas_6day);
if($query_total_lpnas_6day) {
    while ($row = mysqli_fetch_array($query_total_lpnas_6day, MYSQLI_ASSOC)) {
            $count_total_lpnas_6day = $row["counter"];
        }
    }

$sql_total_lpnas_7day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -7 DAY) AND record_status_status='1'";
$query_total_lpnas_7day = mysqli_query($db_conn, $sql_total_lpnas_7day);
if($query_total_lpnas_7day) {
    while ($row = mysqli_fetch_array($query_total_lpnas_7day, MYSQLI_ASSOC)) {
            $count_total_lpnas_7day = $row["counter"];
        }
    }

$sql_total_lpnas_8day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -8 DAY) AND record_status_status='1'";
$query_total_lpnas_8day = mysqli_query($db_conn, $sql_total_lpnas_8day);
if($query_total_lpnas_8day) {
    while ($row = mysqli_fetch_array($query_total_lpnas_8day, MYSQLI_ASSOC)) {
            $count_total_lpnas_8day = $row["counter"];
        }
    }

$sql_total_lpnas_9day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -9 DAY) AND record_status_status='1'";
$query_total_lpnas_9day = mysqli_query($db_conn, $sql_total_lpnas_9day);
if($query_total_lpnas_9day) {
    while ($row = mysqli_fetch_array($query_total_lpnas_9day, MYSQLI_ASSOC)) {
            $count_total_lpnas_9day = $row["counter"];
        }
    }

$sql_total_lpnas_10day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -10 DAY) AND record_status_status='1'";
$query_total_lpnas_10day = mysqli_query($db_conn, $sql_total_lpnas_10day);
if($query_total_lpnas_10day) {
    while ($row = mysqli_fetch_array($query_total_lpnas_10day, MYSQLI_ASSOC)) {
            $count_total_lpnas_10day = $row["counter"];
        }
    }

$sql_total_lpnas_11day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -11 DAY) AND record_status_status='1'";
$query_total_lpnas_11day = mysqli_query($db_conn, $sql_total_lpnas_11day);
if($query_total_lpnas_11day) {
    while ($row = mysqli_fetch_array($query_total_lpnas_11day, MYSQLI_ASSOC)) {
            $count_total_lpnas_11day = $row["counter"];
        }
    }

$sql_total_lpnas_12day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -12 DAY) AND record_status_status='1'";
$query_total_lpnas_12day = mysqli_query($db_conn, $sql_total_lpnas_12day);
if($query_total_lpnas_12day) {
    while ($row = mysqli_fetch_array($query_total_lpnas_12day, MYSQLI_ASSOC)) {
            $count_total_lpnas_12day = $row["counter"];
        }
    }

$sql_total_lpnas_13day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -13 DAY) AND record_status_status='1'";
$query_total_lpnas_13day = mysqli_query($db_conn, $sql_total_lpnas_13day);
if($query_total_lpnas_13day) {
    while ($row = mysqli_fetch_array($query_total_lpnas_13day, MYSQLI_ASSOC)) {
            $count_total_lpnas_13day = $row["counter"];
        }
    }

$sql_total_lpnas_14day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -14 DAY) AND record_status_status='1'";
$query_total_lpnas_14day = mysqli_query($db_conn, $sql_total_lpnas_14day);
if($query_total_lpnas_14day) {
    while ($row = mysqli_fetch_array($query_total_lpnas_14day, MYSQLI_ASSOC)) {
            $count_total_lpnas_14day = $row["counter"];
        }
    }

$sql_total_lpnas_15day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -15 DAY) AND record_status_status='1'";
$query_total_lpnas_15day = mysqli_query($db_conn, $sql_total_lpnas_15day);
if($query_total_lpnas_15day) {
    while ($row = mysqli_fetch_array($query_total_lpnas_15day, MYSQLI_ASSOC)) {
            $count_total_lpnas_15day = $row["counter"];
        }
    }

$totalDay1 = $totalDay1+$count_total_lpnas_1day;
$totalDay2 = $totalDay2+$count_total_lpnas_2day;
$totalDay3 = $totalDay3+$count_total_lpnas_3day;
$totalDay4 = $totalDay4+$count_total_lpnas_4day;
$totalDay5 = $totalDay5+$count_total_lpnas_5day;
$totalDay6 = $totalDay6+$count_total_lpnas_6day;
$totalDay7 = $totalDay7+$count_total_lpnas_7day;
$totalDay8 = $totalDay8+$count_total_lpnas_8day;
$totalDay9 = $totalDay9+$count_total_lpnas_9day;
$totalDay10 = $totalDay10+$count_total_lpnas_10day;
$totalDay11 = $totalDay11+$count_total_lpnas_11day;
$totalDay12 = $totalDay12+$count_total_lpnas_12day;
$totalDay13 = $totalDay13+$count_total_lpnas_13day;
$totalDay14 = $totalDay14+$count_total_lpnas_14day;
$totalDay15 = $totalDay15+$count_total_lpnas_15day;

$ncrTotalDay1 = $ncrTotalDay1+$count_total_lpnas_1day;
$ncrTotalDay2 = $ncrTotalDay2+$count_total_lpnas_2day;
$ncrTotalDay3 = $ncrTotalDay3+$count_total_lpnas_3day;
$ncrTotalDay4 = $ncrTotalDay4+$count_total_lpnas_4day;
$ncrTotalDay5 = $ncrTotalDay5+$count_total_lpnas_5day;
$ncrTotalDay6 = $ncrTotalDay6+$count_total_lpnas_6day;
$ncrTotalDay7 = $ncrTotalDay7+$count_total_lpnas_7day;
$ncrTotalDay8 = $ncrTotalDay8+$count_total_lpnas_8day;
$ncrTotalDay9 = $ncrTotalDay9+$count_total_lpnas_9day;
$ncrTotalDay10 = $ncrTotalDay10+$count_total_lpnas_10day;
$ncrTotalDay11 = $ncrTotalDay11+$count_total_lpnas_11day;
$ncrTotalDay12 = $ncrTotalDay12+$count_total_lpnas_12day;
$ncrTotalDay13 = $ncrTotalDay13+$count_total_lpnas_13day;
$ncrTotalDay14 = $ncrTotalDay14+$count_total_lpnas_14day;
$ncrTotalDay15 = $ncrTotalDay15+$count_total_lpnas_15day;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('H'.$count,$count_total_lpnas_1day)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('I'.$count,$count_total_lpnas_2day)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('J'.$count,$count_total_lpnas_3day)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('K'.$count,$count_total_lpnas_4day)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('L'.$count,$count_total_lpnas_5day)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('M'.$count,$count_total_lpnas_6day)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('N'.$count,$count_total_lpnas_7day)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('O'.$count,$count_total_lpnas_8day)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('P'.$count,$count_total_lpnas_9day)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Q'.$count,$count_total_lpnas_10day)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('R'.$count,$count_total_lpnas_11day)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('S'.$count,$count_total_lpnas_12day)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('T'.$count,$count_total_lpnas_13day)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('U'.$count,$count_total_lpnas_14day)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('V'.$count,$count_total_lpnas_15day)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$count_total_lpnas_delivered = 0;

$sql_total_lpnas_delivered = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_status = 'DELIVERED' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_lpnas_delivered = mysqli_query($db_conn, $sql_total_lpnas_delivered);
if($query_total_lpnas_delivered) {
    while ($row = mysqli_fetch_array($query_total_lpnas_delivered, MYSQLI_ASSOC)) {
            $count_total_lpnas_delivered = $row["counter"];
        }
    }

$totalDelivered = $totalDelivered+$count_total_lpnas_delivered;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('W'.$count,$count_total_lpnas_delivered)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$count_total_lpnas_rts = 0;

$sql_total_lpnas_rts = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_status = 'RTS' AND record_reason_rts != 'OUT OF SCOPE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_lpnas_rts = mysqli_query($db_conn, $sql_total_lpnas_rts);
if($query_total_lpnas_rts) {
    while ($row = mysqli_fetch_array($query_total_lpnas_rts, MYSQLI_ASSOC)) {
            $count_total_lpnas_rts = $row["counter"];
        }
    }

$totalRTS = $totalRTS+$count_total_lpnas_rts;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('X'.$count,$count_total_lpnas_rts)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$count_total_lpnas_oos = 0;

$sql_total_lpnas_oos = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_reason_rts = 'OUT OF SCOPE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_lpnas_oos = mysqli_query($db_conn, $sql_total_lpnas_oos);
if($query_total_lpnas_oos) {
    while ($row = mysqli_fetch_array($query_total_lpnas_oos, MYSQLI_ASSOC)) {
            $count_total_lpnas_oos = $row["counter"];
        }
    }

$totalOOS = $totalOOS+$count_total_lpnas_oos;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Y'.$count,$count_total_lpnas_oos)
        -> getStyle('Y'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y'.$count)
        -> applyFromArray($border);

$count_total_lpnas_balance = 0;

$sql_total_lpnas_balance = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_status = '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_lpnas_balance = mysqli_query($db_conn, $sql_total_lpnas_balance);
if($query_total_lpnas_balance) {
    while ($row = mysqli_fetch_array($query_total_lpnas_balance, MYSQLI_ASSOC)) {
            $count_total_lpnas_balance = $row["counter"];
        }
    }

$totalBalance = $totalBalance+$count_total_lpnas_balance;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Z'.$count,$count_total_lpnas_balance)
        -> getStyle('Z'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

if($count_total_lpnas_balance == 0) {
    $reportStatusLPNAS = "Complete";
} 

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AA'.$count,$reportStatusLPNAS)
        -> getStyle('AA'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AA'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AA'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AB'.$count,'LPNAS')
        -> getStyle('AB'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AB'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AB'.$count)
        -> applyFromArray($border);

$count_total_lpnas1 = 0;

$sql_total_lpnas1 = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_status != '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_lpnas1 = mysqli_query($db_conn, $sql_total_lpnas1);
if($query_total_lpnas1) {
    while ($row = mysqli_fetch_array($query_total_lpnas1, MYSQLI_ASSOC)) {
            $count_total_lpnas1 = $row["counter"];
        }
    }

$count_total_lpnas1 = $count_total_lpnas1-$count_total_lpnas_oos;

$totalVolumeInvoice = $totalVolumeInvoice+$count_total_lpnas1;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AC'.$count,$count_total_lpnas1)
        -> getStyle('AC'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AD'.$count,$lpnasprice)
        -> getStyle('AD'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$amountLPNAS = $lpnasprice*$count_total_lpnas1;

$totalAmount = $totalAmount+$amountLPNAS;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AE'.$count,'Php'.number_format($amountLPNAS))
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);

// 3RD LINE

$count = $count+1;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('A'.$count,strtoupper($sender))
        -> getStyle('A'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('B'.$count,strtoupper($deltype))
        -> getStyle('B'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('B'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('B'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('C'.$count,'MLUPA')
        -> getStyle('C'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('C'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('C'.$count)
        -> applyFromArray($border);

$count_total_mlupa = 0;

$sql_total_mlupa = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mlupa = mysqli_query($db_conn, $sql_total_mlupa);
if($query_total_mlupa) {
    while ($row = mysqli_fetch_array($query_total_mlupa, MYSQLI_ASSOC)) {
            $count_total_mlupa = $row["counter"];
        }
    }

$totalVolume = $totalVolume+$count_total_mlupa;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('G'.$count,$count_total_mlupa)
        -> getStyle('G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G'.$count)
        -> applyFromArray($border);

$count_total_mlupa_1day = 0;
$count_total_mlupa_2day = 0;
$count_total_mlupa_3day = 0;
$count_total_mlupa_4day = 0;
$count_total_mlupa_5day = 0;
$count_total_mlupa_6day = 0;
$count_total_mlupa_7day = 0;
$count_total_mlupa_8day = 0;
$count_total_mlupa_9day = 0;
$count_total_mlupa_10day = 0;
$count_total_mlupa_11day = 0;
$count_total_mlupa_12day = 0;
$count_total_mlupa_13day = 0;
$count_total_mlupa_14day = 0;
$count_total_mlupa_15day = 0;

$sql_total_mlupa_1day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -1 DAY) AND record_status_status='1'";
$query_total_mlupa_1day = mysqli_query($db_conn, $sql_total_mlupa_1day);
if($query_total_mlupa_1day) {
    while ($row = mysqli_fetch_array($query_total_mlupa_1day, MYSQLI_ASSOC)) {
            $count_total_mlupa_1day = $row["counter"];
        }
    }

$sql_total_mlupa_2day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -2 DAY) AND record_status_status='1'";
$query_total_mlupa_2day = mysqli_query($db_conn, $sql_total_mlupa_2day);
if($query_total_mlupa_2day) {
    while ($row = mysqli_fetch_array($query_total_mlupa_2day, MYSQLI_ASSOC)) {
            $count_total_mlupa_2day = $row["counter"];
        }
    }

$sql_total_mlupa_3day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -3 DAY) AND record_status_status='1'";
$query_total_mlupa_3day = mysqli_query($db_conn, $sql_total_mlupa_3day);
if($query_total_mlupa_3day) {
    while ($row = mysqli_fetch_array($query_total_mlupa_3day, MYSQLI_ASSOC)) {
            $count_total_mlupa_3day = $row["counter"];
        }
    }

$sql_total_mlupa_4day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -4 DAY) AND record_status_status='1'";
$query_total_mlupa_4day = mysqli_query($db_conn, $sql_total_mlupa_4day);
if($query_total_mlupa_4day) {
    while ($row = mysqli_fetch_array($query_total_mlupa_4day, MYSQLI_ASSOC)) {
            $count_total_mlupa_4day = $row["counter"];
        }
    }

$sql_total_mlupa_5day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -5 DAY) AND record_status_status='1'";
$query_total_mlupa_5day = mysqli_query($db_conn, $sql_total_mlupa_5day);
if($query_total_mlupa_5day) {
    while ($row = mysqli_fetch_array($query_total_mlupa_5day, MYSQLI_ASSOC)) {
            $count_total_mlupa_5day = $row["counter"];
        }
    }

$sql_total_mlupa_6day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -6 DAY) AND record_status_status='1'";
$query_total_mlupa_6day = mysqli_query($db_conn, $sql_total_mlupa_6day);
if($query_total_mlupa_6day) {
    while ($row = mysqli_fetch_array($query_total_mlupa_6day, MYSQLI_ASSOC)) {
            $count_total_mlupa_6day = $row["counter"];
        }
    }

$sql_total_mlupa_7day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -7 DAY) AND record_status_status='1'";
$query_total_mlupa_7day = mysqli_query($db_conn, $sql_total_mlupa_7day);
if($query_total_mlupa_7day) {
    while ($row = mysqli_fetch_array($query_total_mlupa_7day, MYSQLI_ASSOC)) {
            $count_total_mlupa_7day = $row["counter"];
        }
    }

$sql_total_mlupa_8day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -8 DAY) AND record_status_status='1'";
$query_total_mlupa_8day = mysqli_query($db_conn, $sql_total_mlupa_8day);
if($query_total_mlupa_8day) {
    while ($row = mysqli_fetch_array($query_total_mlupa_8day, MYSQLI_ASSOC)) {
            $count_total_mlupa_8day = $row["counter"];
        }
    }

$sql_total_mlupa_9day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -9 DAY) AND record_status_status='1'";
$query_total_mlupa_9day = mysqli_query($db_conn, $sql_total_mlupa_9day);
if($query_total_mlupa_9day) {
    while ($row = mysqli_fetch_array($query_total_mlupa_9day, MYSQLI_ASSOC)) {
            $count_total_mlupa_9day = $row["counter"];
        }
    }

$sql_total_mlupa_10day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -10 DAY) AND record_status_status='1'";
$query_total_mlupa_10day = mysqli_query($db_conn, $sql_total_mlupa_10day);
if($query_total_mlupa_10day) {
    while ($row = mysqli_fetch_array($query_total_mlupa_10day, MYSQLI_ASSOC)) {
            $count_total_mlupa_10day = $row["counter"];
        }
    }

$sql_total_mlupa_11day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -11 DAY) AND record_status_status='1'";
$query_total_mlupa_11day = mysqli_query($db_conn, $sql_total_mlupa_11day);
if($query_total_mlupa_11day) {
    while ($row = mysqli_fetch_array($query_total_mlupa_11day, MYSQLI_ASSOC)) {
            $count_total_mlupa_11day = $row["counter"];
        }
    }

$sql_total_mlupa_12day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -12 DAY) AND record_status_status='1'";
$query_total_mlupa_12day = mysqli_query($db_conn, $sql_total_mlupa_12day);
if($query_total_mlupa_12day) {
    while ($row = mysqli_fetch_array($query_total_mlupa_12day, MYSQLI_ASSOC)) {
            $count_total_mlupa_12day = $row["counter"];
        }
    }

$sql_total_mlupa_13day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -13 DAY) AND record_status_status='1'";
$query_total_mlupa_13day = mysqli_query($db_conn, $sql_total_mlupa_13day);
if($query_total_mlupa_13day) {
    while ($row = mysqli_fetch_array($query_total_mlupa_13day, MYSQLI_ASSOC)) {
            $count_total_mlupa_13day = $row["counter"];
        }
    }

$sql_total_mlupa_14day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -14 DAY) AND record_status_status='1'";
$query_total_mlupa_14day = mysqli_query($db_conn, $sql_total_mlupa_14day);
if($query_total_mlupa_14day) {
    while ($row = mysqli_fetch_array($query_total_mlupa_14day, MYSQLI_ASSOC)) {
            $count_total_mlupa_14day = $row["counter"];
        }
    }

$sql_total_mlupa_15day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -15 DAY) AND record_status_status='1'";
$query_total_mlupa_15day = mysqli_query($db_conn, $sql_total_mlupa_15day);
if($query_total_mlupa_15day) {
    while ($row = mysqli_fetch_array($query_total_mlupa_15day, MYSQLI_ASSOC)) {
            $count_total_mlupa_15day = $row["counter"];
        }
    }

$totalDay1 = $totalDay1+$count_total_mlupa_1day;
$totalDay2 = $totalDay2+$count_total_mlupa_2day;
$totalDay3 = $totalDay3+$count_total_mlupa_3day;
$totalDay4 = $totalDay4+$count_total_mlupa_4day;
$totalDay5 = $totalDay5+$count_total_mlupa_5day;
$totalDay6 = $totalDay6+$count_total_mlupa_6day;
$totalDay7 = $totalDay7+$count_total_mlupa_7day;
$totalDay8 = $totalDay8+$count_total_mlupa_8day;
$totalDay9 = $totalDay9+$count_total_mlupa_9day;
$totalDay10 = $totalDay10+$count_total_mlupa_10day;
$totalDay11 = $totalDay11+$count_total_mlupa_11day;
$totalDay12 = $totalDay12+$count_total_mlupa_12day;
$totalDay13 = $totalDay13+$count_total_mlupa_13day;
$totalDay14 = $totalDay14+$count_total_mlupa_14day;
$totalDay15 = $totalDay15+$count_total_mlupa_15day;

$ncrTotalDay1 = $ncrTotalDay1+$count_total_mlupa_1day;
$ncrTotalDay2 = $ncrTotalDay2+$count_total_mlupa_2day;
$ncrTotalDay3 = $ncrTotalDay3+$count_total_mlupa_3day;
$ncrTotalDay4 = $ncrTotalDay4+$count_total_mlupa_4day;
$ncrTotalDay5 = $ncrTotalDay5+$count_total_mlupa_5day;
$ncrTotalDay6 = $ncrTotalDay6+$count_total_mlupa_6day;
$ncrTotalDay7 = $ncrTotalDay7+$count_total_mlupa_7day;
$ncrTotalDay8 = $ncrTotalDay8+$count_total_mlupa_8day;
$ncrTotalDay9 = $ncrTotalDay9+$count_total_mlupa_9day;
$ncrTotalDay10 = $ncrTotalDay10+$count_total_mlupa_10day;
$ncrTotalDay11 = $ncrTotalDay11+$count_total_mlupa_11day;
$ncrTotalDay12 = $ncrTotalDay12+$count_total_mlupa_12day;
$ncrTotalDay13 = $ncrTotalDay13+$count_total_mlupa_13day;
$ncrTotalDay14 = $ncrTotalDay14+$count_total_mlupa_14day;
$ncrTotalDay15 = $ncrTotalDay15+$count_total_mlupa_15day;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('H'.$count,$count_total_mlupa_1day)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('I'.$count,$count_total_mlupa_2day)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('J'.$count,$count_total_mlupa_3day)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('K'.$count,$count_total_mlupa_4day)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('L'.$count,$count_total_mlupa_5day)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('M'.$count,$count_total_mlupa_6day)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('N'.$count,$count_total_mlupa_7day)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('O'.$count,$count_total_mlupa_8day)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('P'.$count,$count_total_mlupa_9day)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Q'.$count,$count_total_mlupa_10day)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('R'.$count,$count_total_mlupa_11day)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('S'.$count,$count_total_mlupa_12day)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('T'.$count,$count_total_mlupa_13day)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('U'.$count,$count_total_mlupa_14day)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('V'.$count,$count_total_mlupa_15day)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$count_total_mlupa_delivered = 0;

$sql_total_mlupa_delivered = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_status = 'DELIVERED' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mlupa_delivered = mysqli_query($db_conn, $sql_total_mlupa_delivered);
if($query_total_mlupa_delivered) {
    while ($row = mysqli_fetch_array($query_total_mlupa_delivered, MYSQLI_ASSOC)) {
            $count_total_mlupa_delivered = $row["counter"];
        }
    }

$totalDelivered = $totalDelivered+$count_total_mlupa_delivered;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('W'.$count,$count_total_mlupa_delivered)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$count_total_mlupa_rts = 0;

$sql_total_mlupa_rts = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_status = 'RTS' AND record_reason_rts != 'OUT OF SCOPE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mlupa_rts = mysqli_query($db_conn, $sql_total_mlupa_rts);
if($query_total_mlupa_rts) {
    while ($row = mysqli_fetch_array($query_total_mlupa_rts, MYSQLI_ASSOC)) {
            $count_total_mlupa_rts = $row["counter"];
        }
    }

$totalRTS = $totalRTS+$count_total_mlupa_rts;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('X'.$count,$count_total_mlupa_rts)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$count_total_mlupa_oos = 0;

$sql_total_mlupa_oos = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_reason_rts = 'OUT OF SCOPE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mlupa_oos = mysqli_query($db_conn, $sql_total_mlupa_oos);
if($query_total_mlupa_oos) {
    while ($row = mysqli_fetch_array($query_total_mlupa_oos, MYSQLI_ASSOC)) {
            $count_total_mlupa_oos = $row["counter"];
        }
    }

$totalOOS = $totalOOS+$count_total_mlupa_oos;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Y'.$count,$count_total_mlupa_oos)
        -> getStyle('Y'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y'.$count)
        -> applyFromArray($border);

$count_total_mlupa_balance = 0;

$sql_total_mlupa_balance = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_status = '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mlupa_balance = mysqli_query($db_conn, $sql_total_mlupa_balance);
if($query_total_mlupa_balance) {
    while ($row = mysqli_fetch_array($query_total_mlupa_balance, MYSQLI_ASSOC)) {
            $count_total_mlupa_balance = $row["counter"];
        }
    }

$totalBalance = $totalBalance+$count_total_mlupa_balance;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Z'.$count,$count_total_mlupa_balance)
        -> getStyle('Z'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

if($count_total_mlupa_balance == 0) {
    $reportStatusMLUPA = "Complete";
}

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AA'.$count,$reportStatusMLUPA)
        -> getStyle('AA'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AA'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AA'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AB'.$count,'MLUPA')
        -> getStyle('AB'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AB'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AB'.$count)
        -> applyFromArray($border);

$count_total_mlupa1 = 0;

$sql_total_mlupa1 = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MUNTINLUPA' AND record_status != '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mlupa1 = mysqli_query($db_conn, $sql_total_mlupa1);
if($query_total_mlupa1) {
    while ($row = mysqli_fetch_array($query_total_mlupa1, MYSQLI_ASSOC)) {
            $count_total_mlupa1 = $row["counter"];
        }
    }

$count_total_mlupa1 = $count_total_mlupa1-$count_total_mlupa_oos;

$totalVolumeInvoice = $totalVolumeInvoice+$count_total_mlupa1;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AC'.$count,$count_total_mlupa1)
        -> getStyle('AC'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AD'.$count,$mlupaprice)
        -> getStyle('AD'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$amountMLUPA = $mlupaprice*$count_total_mlupa1;

$totalAmount = $totalAmount+$amountMLUPA;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AE'.$count,'Php'.number_format($amountMLUPA))
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);

// 4TH LINE

$count = $count+1;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('A'.$count,strtoupper($sender))
        -> getStyle('A'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('B'.$count,strtoupper($deltype))
        -> getStyle('B'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('B'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('B'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('C'.$count,'MKNA')
        -> getStyle('C'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('C'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('C'.$count)
        -> applyFromArray($border);

$count_total_mkna = 0;

$sql_total_mkna = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mkna = mysqli_query($db_conn, $sql_total_mkna);
if($query_total_mkna) {
    while ($row = mysqli_fetch_array($query_total_mkna, MYSQLI_ASSOC)) {
            $count_total_mkna = $row["counter"];
        }
    }

$totalVolume = $totalVolume+$count_total_mkna;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('G'.$count,$count_total_mkna)
        -> getStyle('G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G'.$count)
        -> applyFromArray($border);

$count_total_mkna_1day = 0;
$count_total_mkna_2day = 0;
$count_total_mkna_3day = 0;
$count_total_mkna_4day = 0;
$count_total_mkna_5day = 0;
$count_total_mkna_6day = 0;
$count_total_mkna_7day = 0;
$count_total_mkna_8day = 0;
$count_total_mkna_9day = 0;
$count_total_mkna_10day = 0;
$count_total_mkna_11day = 0;
$count_total_mkna_12day = 0;
$count_total_mkna_13day = 0;
$count_total_mkna_14day = 0;
$count_total_mkna_15day = 0;

$sql_total_mkna_1day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -1 DAY) AND record_status_status='1'";
$query_total_mkna_1day = mysqli_query($db_conn, $sql_total_mkna_1day);
if($query_total_mkna_1day) {
    while ($row = mysqli_fetch_array($query_total_mkna_1day, MYSQLI_ASSOC)) {
            $count_total_mkna_1day = $row["counter"];
        }
    }

$sql_total_mkna_2day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -2 DAY) AND record_status_status='1'";
$query_total_mkna_2day = mysqli_query($db_conn, $sql_total_mkna_2day);
if($query_total_mkna_2day) {
    while ($row = mysqli_fetch_array($query_total_mkna_2day, MYSQLI_ASSOC)) {
            $count_total_mkna_2day = $row["counter"];
        }
    }

$sql_total_mkna_3day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -3 DAY) AND record_status_status='1'";
$query_total_mkna_3day = mysqli_query($db_conn, $sql_total_mkna_3day);
if($query_total_mkna_3day) {
    while ($row = mysqli_fetch_array($query_total_mkna_3day, MYSQLI_ASSOC)) {
            $count_total_mkna_3day = $row["counter"];
        }
    }

$sql_total_mkna_4day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -4 DAY) AND record_status_status='1'";
$query_total_mkna_4day = mysqli_query($db_conn, $sql_total_mkna_4day);
if($query_total_mkna_4day) {
    while ($row = mysqli_fetch_array($query_total_mkna_4day, MYSQLI_ASSOC)) {
            $count_total_mkna_4day = $row["counter"];
        }
    }

$sql_total_mkna_5day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -5 DAY) AND record_status_status='1'";
$query_total_mkna_5day = mysqli_query($db_conn, $sql_total_mkna_5day);
if($query_total_mkna_5day) {
    while ($row = mysqli_fetch_array($query_total_mkna_5day, MYSQLI_ASSOC)) {
            $count_total_mkna_5day = $row["counter"];
        }
    }

$sql_total_mkna_6day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -6 DAY) AND record_status_status='1'";
$query_total_mkna_6day = mysqli_query($db_conn, $sql_total_mkna_6day);
if($query_total_mkna_6day) {
    while ($row = mysqli_fetch_array($query_total_mkna_6day, MYSQLI_ASSOC)) {
            $count_total_mkna_6day = $row["counter"];
        }
    }

$sql_total_mkna_7day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -7 DAY) AND record_status_status='1'";
$query_total_mkna_7day = mysqli_query($db_conn, $sql_total_mkna_7day);
if($query_total_mkna_7day) {
    while ($row = mysqli_fetch_array($query_total_mkna_7day, MYSQLI_ASSOC)) {
            $count_total_mkna_7day = $row["counter"];
        }
    }

$sql_total_mkna_8day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -8 DAY) AND record_status_status='1'";
$query_total_mkna_8day = mysqli_query($db_conn, $sql_total_mkna_8day);
if($query_total_mkna_8day) {
    while ($row = mysqli_fetch_array($query_total_mkna_8day, MYSQLI_ASSOC)) {
            $count_total_mkna_8day = $row["counter"];
        }
    }

$sql_total_mkna_9day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -9 DAY) AND record_status_status='1'";
$query_total_mkna_9day = mysqli_query($db_conn, $sql_total_mkna_9day);
if($query_total_mkna_9day) {
    while ($row = mysqli_fetch_array($query_total_mkna_9day, MYSQLI_ASSOC)) {
            $count_total_mkna_9day = $row["counter"];
        }
    }

$sql_total_mkna_10day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -10 DAY) AND record_status_status='1'";
$query_total_mkna_10day = mysqli_query($db_conn, $sql_total_mkna_10day);
if($query_total_mkna_10day) {
    while ($row = mysqli_fetch_array($query_total_mkna_10day, MYSQLI_ASSOC)) {
            $count_total_mkna_10day = $row["counter"];
        }
    }

$sql_total_mkna_11day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -11 DAY) AND record_status_status='1'";
$query_total_mkna_11day = mysqli_query($db_conn, $sql_total_mkna_11day);
if($query_total_mkna_11day) {
    while ($row = mysqli_fetch_array($query_total_mkna_11day, MYSQLI_ASSOC)) {
            $count_total_mkna_11day = $row["counter"];
        }
    }

$sql_total_mkna_12day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -12 DAY) AND record_status_status='1'";
$query_total_mkna_12day = mysqli_query($db_conn, $sql_total_mkna_12day);
if($query_total_mkna_12day) {
    while ($row = mysqli_fetch_array($query_total_mkna_12day, MYSQLI_ASSOC)) {
            $count_total_mkna_12day = $row["counter"];
        }
    }

$sql_total_mkna_13day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -13 DAY) AND record_status_status='1'";
$query_total_mkna_13day = mysqli_query($db_conn, $sql_total_mkna_13day);
if($query_total_mkna_13day) {
    while ($row = mysqli_fetch_array($query_total_mkna_13day, MYSQLI_ASSOC)) {
            $count_total_mkna_13day = $row["counter"];
        }
    }

$sql_total_mkna_14day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -14 DAY) AND record_status_status='1'";
$query_total_mkna_14day = mysqli_query($db_conn, $sql_total_mkna_14day);
if($query_total_mkna_14day) {
    while ($row = mysqli_fetch_array($query_total_mkna_14day, MYSQLI_ASSOC)) {
            $count_total_mkna_14day = $row["counter"];
        }
    }

$sql_total_mkna_15day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -15 DAY) AND record_status_status='1'";
$query_total_mkna_15day = mysqli_query($db_conn, $sql_total_mkna_15day);
if($query_total_mkna_15day) {
    while ($row = mysqli_fetch_array($query_total_mkna_15day, MYSQLI_ASSOC)) {
            $count_total_mkna_15day = $row["counter"];
        }
    }

$totalDay1 = $totalDay1+$count_total_mkna_1day;
$totalDay2 = $totalDay2+$count_total_mkna_2day;
$totalDay3 = $totalDay3+$count_total_mkna_3day;
$totalDay4 = $totalDay4+$count_total_mkna_4day;
$totalDay5 = $totalDay5+$count_total_mkna_5day;
$totalDay6 = $totalDay6+$count_total_mkna_6day;
$totalDay7 = $totalDay7+$count_total_mkna_7day;
$totalDay8 = $totalDay8+$count_total_mkna_8day;
$totalDay9 = $totalDay9+$count_total_mkna_9day;
$totalDay10 = $totalDay10+$count_total_mkna_10day;
$totalDay11 = $totalDay11+$count_total_mkna_11day;
$totalDay12 = $totalDay12+$count_total_mkna_12day;
$totalDay13 = $totalDay13+$count_total_mkna_13day;
$totalDay14 = $totalDay14+$count_total_mkna_14day;
$totalDay15 = $totalDay15+$count_total_mkna_15day;

$ncrTotalDay1 = $ncrTotalDay1+$count_total_mkna_1day;
$ncrTotalDay2 = $ncrTotalDay2+$count_total_mkna_2day;
$ncrTotalDay3 = $ncrTotalDay3+$count_total_mkna_3day;
$ncrTotalDay4 = $ncrTotalDay4+$count_total_mkna_4day;
$ncrTotalDay5 = $ncrTotalDay5+$count_total_mkna_5day;
$ncrTotalDay6 = $ncrTotalDay6+$count_total_mkna_6day;
$ncrTotalDay7 = $ncrTotalDay7+$count_total_mkna_7day;
$ncrTotalDay8 = $ncrTotalDay8+$count_total_mkna_8day;
$ncrTotalDay9 = $ncrTotalDay9+$count_total_mkna_9day;
$ncrTotalDay10 = $ncrTotalDay10+$count_total_mkna_10day;
$ncrTotalDay11 = $ncrTotalDay11+$count_total_mkna_11day;
$ncrTotalDay12 = $ncrTotalDay12+$count_total_mkna_12day;
$ncrTotalDay13 = $ncrTotalDay13+$count_total_mkna_13day;
$ncrTotalDay14 = $ncrTotalDay14+$count_total_mkna_14day;
$ncrTotalDay15 = $ncrTotalDay15+$count_total_mkna_15day;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('H'.$count,$count_total_mkna_1day)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('I'.$count,$count_total_mkna_2day)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('J'.$count,$count_total_mkna_3day)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('K'.$count,$count_total_mkna_4day)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('L'.$count,$count_total_mkna_5day)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('M'.$count,$count_total_mkna_6day)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('N'.$count,$count_total_mkna_7day)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('O'.$count,$count_total_mkna_8day)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('P'.$count,$count_total_mkna_9day)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Q'.$count,$count_total_mkna_10day)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('R'.$count,$count_total_mkna_11day)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('S'.$count,$count_total_mkna_12day)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('T'.$count,$count_total_mkna_13day)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('U'.$count,$count_total_mkna_14day)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('V'.$count,$count_total_mkna_15day)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$count_total_mkna_delivered = 0;

$sql_total_mkna_delivered = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_status = 'DELIVERED' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mkna_delivered = mysqli_query($db_conn, $sql_total_mkna_delivered);
if($query_total_mkna_delivered) {
    while ($row = mysqli_fetch_array($query_total_mkna_delivered, MYSQLI_ASSOC)) {
            $count_total_mkna_delivered = $row["counter"];
        }
    }

$totalDelivered = $totalDelivered+$count_total_mkna_delivered;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('W'.$count,$count_total_mkna_delivered)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$count_total_mkna_rts = 0;

$sql_total_mkna_rts = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_status = 'RTS' AND record_reason_rts != 'OUT OF SCOPE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mkna_rts = mysqli_query($db_conn, $sql_total_mkna_rts);
if($query_total_mkna_rts) {
    while ($row = mysqli_fetch_array($query_total_mkna_rts, MYSQLI_ASSOC)) {
            $count_total_mkna_rts = $row["counter"];
        }
    }

$totalRTS = $totalRTS+$count_total_mkna_rts;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('X'.$count,$count_total_mkna_rts)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$count_total_mkna_oos = 0;

$sql_total_mkna_oos = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_reason_rts = 'OUT OF SCOPE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mkna_oos = mysqli_query($db_conn, $sql_total_mkna_oos);
if($query_total_mkna_oos) {
    while ($row = mysqli_fetch_array($query_total_mkna_oos, MYSQLI_ASSOC)) {
            $count_total_mkna_oos = $row["counter"];
        }
    }

$totalOOS = $totalOOS+$count_total_mkna_oos;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Y'.$count,$count_total_mkna_oos)
        -> getStyle('Y'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y'.$count)
        -> applyFromArray($border);

$count_total_mkna_balance = 0;

$sql_total_mkna_balance = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_status = '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mkna_balance = mysqli_query($db_conn, $sql_total_mkna_balance);
if($query_total_mkna_balance) {
    while ($row = mysqli_fetch_array($query_total_mkna_balance, MYSQLI_ASSOC)) {
            $count_total_mkna_balance = $row["counter"];
        }
    }

$totalBalance = $totalBalance+$count_total_mkna_balance;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Z'.$count,$count_total_mkna_balance)
        -> getStyle('Z'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

if($count_total_mkna_balance == 0) {
    $reportStatusMKNA = "Complete";
}

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AA'.$count,$reportStatusMKNA)
        -> getStyle('AA'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AA'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AA'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AB'.$count,'MKNA')
        -> getStyle('AB'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AB'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AB'.$count)
        -> applyFromArray($border);

$count_total_mkna1 = 0;

$sql_total_mkna1 = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'MARIKINA' AND record_status != '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mkna1 = mysqli_query($db_conn, $sql_total_mkna1);
if($query_total_mkna1) {
    while ($row = mysqli_fetch_array($query_total_mkna1, MYSQLI_ASSOC)) {
            $count_total_mkna1 = $row["counter"];
        }
    }

$count_total_mkna1 = $count_total_mkna1-$count_total_mkna_oos;

$totalVolumeInvoice = $totalVolumeInvoice+$count_total_mkna1;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AC'.$count,$count_total_mkna1)
        -> getStyle('AC'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AD'.$count,$mknaprice)
        -> getStyle('AD'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$amountMKNA = $mknaprice*$count_total_mkna1;

$totalAmount = $totalAmount+$amountMKNA;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AE'.$count,'Php'.number_format($amountMKNA))
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);

//5TH LINE

$count = $count+1;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('A'.$count,strtoupper($sender))
        -> getStyle('A'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('B'.$count,strtoupper($deltype))
        -> getStyle('B'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('B'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('B'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('C'.$count,'RIZAL')
        -> getStyle('C'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('C'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('C'.$count)
        -> applyFromArray($border);

$count_total_rizal = 0;

$sql_total_rizal = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_rizal = mysqli_query($db_conn, $sql_total_rizal);
if($query_total_rizal) {
    while ($row = mysqli_fetch_array($query_total_rizal, MYSQLI_ASSOC)) {
            $count_total_rizal = $row["counter"];
        }
    }

$totalVolume = $totalVolume+$count_total_rizal;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('G'.$count,$count_total_rizal)
        -> getStyle('G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G'.$count)
        -> applyFromArray($border);

$count_total_rizal_1day = 0;
$count_total_rizal_2day = 0;
$count_total_rizal_3day = 0;
$count_total_rizal_4day = 0;
$count_total_rizal_5day = 0;
$count_total_rizal_6day = 0;
$count_total_rizal_7day = 0;
$count_total_rizal_8day = 0;
$count_total_rizal_9day = 0;
$count_total_rizal_10day = 0;
$count_total_rizal_11day = 0;
$count_total_rizal_12day = 0;
$count_total_rizal_13day = 0;
$count_total_rizal_14day = 0;
$count_total_rizal_15day = 0;

$sql_total_rizal_1day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -1 DAY) AND record_status_status='1'";
$query_total_rizal_1day = mysqli_query($db_conn, $sql_total_rizal_1day);
if($query_total_rizal_1day) {
    while ($row = mysqli_fetch_array($query_total_rizal_1day, MYSQLI_ASSOC)) {
            $count_total_rizal_1day = $row["counter"];
        }
    }

$sql_total_rizal_2day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -2 DAY) AND record_status_status='1'";
$query_total_rizal_2day = mysqli_query($db_conn, $sql_total_rizal_2day);
if($query_total_rizal_2day) {
    while ($row = mysqli_fetch_array($query_total_rizal_2day, MYSQLI_ASSOC)) {
            $count_total_rizal_2day = $row["counter"];
        }
    }

$sql_total_rizal_3day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -3 DAY) AND record_status_status='1'";
$query_total_rizal_3day = mysqli_query($db_conn, $sql_total_rizal_3day);
if($query_total_rizal_3day) {
    while ($row = mysqli_fetch_array($query_total_rizal_3day, MYSQLI_ASSOC)) {
            $count_total_rizal_3day = $row["counter"];
        }
    }

$sql_total_rizal_4day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -4 DAY) AND record_status_status='1'";
$query_total_rizal_4day = mysqli_query($db_conn, $sql_total_rizal_4day);
if($query_total_rizal_4day) {
    while ($row = mysqli_fetch_array($query_total_rizal_4day, MYSQLI_ASSOC)) {
            $count_total_rizal_4day = $row["counter"];
        }
    }

$sql_total_rizal_5day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -5 DAY) AND record_status_status='1'";
$query_total_rizal_5day = mysqli_query($db_conn, $sql_total_rizal_5day);
if($query_total_rizal_5day) {
    while ($row = mysqli_fetch_array($query_total_rizal_5day, MYSQLI_ASSOC)) {
            $count_total_rizal_5day = $row["counter"];
        }
    }

$sql_total_rizal_6day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -6 DAY) AND record_status_status='1'";
$query_total_rizal_6day = mysqli_query($db_conn, $sql_total_rizal_6day);
if($query_total_rizal_6day) {
    while ($row = mysqli_fetch_array($query_total_rizal_6day, MYSQLI_ASSOC)) {
            $count_total_rizal_6day = $row["counter"];
        }
    }

$sql_total_rizal_7day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -7 DAY) AND record_status_status='1'";
$query_total_rizal_7day = mysqli_query($db_conn, $sql_total_rizal_7day);
if($query_total_rizal_7day) {
    while ($row = mysqli_fetch_array($query_total_rizal_7day, MYSQLI_ASSOC)) {
            $count_total_rizal_7day = $row["counter"];
        }
    }

$sql_total_rizal_8day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -8 DAY) AND record_status_status='1'";
$query_total_rizal_8day = mysqli_query($db_conn, $sql_total_rizal_8day);
if($query_total_rizal_8day) {
    while ($row = mysqli_fetch_array($query_total_rizal_8day, MYSQLI_ASSOC)) {
            $count_total_rizal_8day = $row["counter"];
        }
    }

$sql_total_rizal_9day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -9 DAY) AND record_status_status='1'";
$query_total_rizal_9day = mysqli_query($db_conn, $sql_total_rizal_9day);
if($query_total_rizal_9day) {
    while ($row = mysqli_fetch_array($query_total_rizal_9day, MYSQLI_ASSOC)) {
            $count_total_rizal_9day = $row["counter"];
        }
    }

$sql_total_rizal_10day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -10 DAY) AND record_status_status='1'";
$query_total_rizal_10day = mysqli_query($db_conn, $sql_total_rizal_10day);
if($query_total_rizal_10day) {
    while ($row = mysqli_fetch_array($query_total_rizal_10day, MYSQLI_ASSOC)) {
            $count_total_rizal_10day = $row["counter"];
        }
    }

$sql_total_rizal_11day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -11 DAY) AND record_status_status='1'";
$query_total_rizal_11day = mysqli_query($db_conn, $sql_total_rizal_11day);
if($query_total_rizal_11day) {
    while ($row = mysqli_fetch_array($query_total_rizal_11day, MYSQLI_ASSOC)) {
            $count_total_rizal_11day = $row["counter"];
        }
    }

$sql_total_rizal_12day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -12 DAY) AND record_status_status='1'";
$query_total_rizal_12day = mysqli_query($db_conn, $sql_total_rizal_12day);
if($query_total_rizal_12day) {
    while ($row = mysqli_fetch_array($query_total_rizal_12day, MYSQLI_ASSOC)) {
            $count_total_rizal_12day = $row["counter"];
        }
    }

$sql_total_rizal_13day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -13 DAY) AND record_status_status='1'";
$query_total_rizal_13day = mysqli_query($db_conn, $sql_total_rizal_13day);
if($query_total_rizal_13day) {
    while ($row = mysqli_fetch_array($query_total_rizal_13day, MYSQLI_ASSOC)) {
            $count_total_rizal_13day = $row["counter"];
        }
    }

$sql_total_rizal_14day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -14 DAY) AND record_status_status='1'";
$query_total_rizal_14day = mysqli_query($db_conn, $sql_total_rizal_14day);
if($query_total_rizal_14day) {
    while ($row = mysqli_fetch_array($query_total_rizal_14day, MYSQLI_ASSOC)) {
            $count_total_rizal_14day = $row["counter"];
        }
    }

$sql_total_rizal_15day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -15 DAY) AND record_status_status='1'";
$query_total_rizal_15day = mysqli_query($db_conn, $sql_total_rizal_15day);
if($query_total_rizal_15day) {
    while ($row = mysqli_fetch_array($query_total_rizal_15day, MYSQLI_ASSOC)) {
            $count_total_rizal_15day = $row["counter"];
        }
    }

$totalDay1 = $totalDay1+$count_total_rizal_1day;
$totalDay2 = $totalDay2+$count_total_rizal_2day;
$totalDay3 = $totalDay3+$count_total_rizal_3day;
$totalDay4 = $totalDay4+$count_total_rizal_4day;
$totalDay5 = $totalDay5+$count_total_rizal_5day;
$totalDay6 = $totalDay6+$count_total_rizal_6day;
$totalDay7 = $totalDay7+$count_total_rizal_7day;
$totalDay8 = $totalDay8+$count_total_rizal_8day;
$totalDay9 = $totalDay9+$count_total_rizal_9day;
$totalDay10 = $totalDay10+$count_total_rizal_10day;
$totalDay11 = $totalDay11+$count_total_rizal_11day;
$totalDay12 = $totalDay12+$count_total_rizal_12day;
$totalDay13 = $totalDay13+$count_total_rizal_13day;
$totalDay14 = $totalDay14+$count_total_rizal_14day;
$totalDay15 = $totalDay15+$count_total_rizal_15day;

$provTotalDay1 = $provTotalDay1+$count_total_rizal_1day;
$provTotalDay2 = $provTotalDay2+$count_total_rizal_2day;
$provTotalDay3 = $provTotalDay3+$count_total_rizal_3day;
$provTotalDay4 = $provTotalDay4+$count_total_rizal_4day;
$provTotalDay5 = $provTotalDay5+$count_total_rizal_5day;
$provTotalDay6 = $provTotalDay6+$count_total_rizal_6day;
$provTotalDay7 = $provTotalDay7+$count_total_rizal_7day;
$provTotalDay8 = $provTotalDay8+$count_total_rizal_8day;
$provTotalDay9 = $provTotalDay9+$count_total_rizal_9day;
$provTotalDay10 = $provTotalDay10+$count_total_rizal_10day;
$provTotalDay11 = $provTotalDay11+$count_total_rizal_11day;
$provTotalDay12 = $provTotalDay12+$count_total_rizal_12day;
$provTotalDay13 = $provTotalDay13+$count_total_rizal_13day;
$provTotalDay14 = $provTotalDay14+$count_total_rizal_14day;
$provTotalDay15 = $provTotalDay15+$count_total_rizal_15day;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('H'.$count,$count_total_rizal_1day)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('I'.$count,$count_total_rizal_2day)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('J'.$count,$count_total_rizal_3day)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('K'.$count,$count_total_rizal_4day)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('L'.$count,$count_total_rizal_5day)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('M'.$count,$count_total_rizal_6day)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('N'.$count,$count_total_rizal_7day)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('O'.$count,$count_total_rizal_8day)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('P'.$count,$count_total_rizal_9day)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Q'.$count,$count_total_rizal_10day)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('R'.$count,$count_total_rizal_11day)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('S'.$count,$count_total_rizal_12day)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('T'.$count,$count_total_rizal_13day)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('U'.$count,$count_total_rizal_14day)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('V'.$count,$count_total_rizal_15day)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$count_total_rizal_delivered = 0;

$sql_total_rizal_delivered = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_status = 'DELIVERED' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_rizal_delivered = mysqli_query($db_conn, $sql_total_rizal_delivered);
if($query_total_rizal_delivered) {
    while ($row = mysqli_fetch_array($query_total_rizal_delivered, MYSQLI_ASSOC)) {
            $count_total_rizal_delivered = $row["counter"];
        }
    }

$totalDelivered = $totalDelivered+$count_total_rizal_delivered;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('W'.$count,$count_total_rizal_delivered)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$count_total_rizal_rts = 0;

$sql_total_rizal_rts = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_status = 'RTS' AND record_reason_rts != 'OUT OF SCOPE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_rizal_rts = mysqli_query($db_conn, $sql_total_rizal_rts);
if($query_total_rizal_rts) {
    while ($row = mysqli_fetch_array($query_total_rizal_rts, MYSQLI_ASSOC)) {
            $count_total_rizal_rts = $row["counter"];
        }
    }

$totalRTS = $totalRTS+$count_total_rizal_rts;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('X'.$count,$count_total_rizal_rts)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$count_total_rizal_oos = 0;

$sql_total_rizal_oos = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_reason_rts = 'OUT OF SCOPE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_rizal_oos = mysqli_query($db_conn, $sql_total_rizal_oos);
if($query_total_rizal_oos) {
    while ($row = mysqli_fetch_array($query_total_rizal_oos, MYSQLI_ASSOC)) {
            $count_total_rizal_oos = $row["counter"];
        }
    }

$totalOOS = $totalOOS+$count_total_rizal_oos;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Y'.$count,$count_total_rizal_oos)
        -> getStyle('Y'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y'.$count)
        -> applyFromArray($border);

$count_total_rizal_balance = 0;

$sql_total_rizal_balance = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_status = '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_rizal_balance = mysqli_query($db_conn, $sql_total_rizal_balance);
if($query_total_rizal_balance) {
    while ($row = mysqli_fetch_array($query_total_rizal_balance, MYSQLI_ASSOC)) {
            $count_total_rizal_balance = $row["counter"];
        }
    }

$totalBalance = $totalBalance+$count_total_rizal_balance;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Z'.$count,$count_total_rizal_balance)
        -> getStyle('Z'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

if($count_total_rizal_balance == 0) {
    $reportStatusRIZAL = "Complete";
}

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AA'.$count,$reportStatusRIZAL)
        -> getStyle('AA'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AA'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AA'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AB'.$count,'RIZAL')
        -> getStyle('AB'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AB'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AB'.$count)
        -> applyFromArray($border);

$count_total_rizal1 = 0;        

$sql_total_rizal1 = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'RIZAL' AND record_status != '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_rizal1 = mysqli_query($db_conn, $sql_total_rizal1);
if($query_total_rizal1) {
    while ($row = mysqli_fetch_array($query_total_rizal1, MYSQLI_ASSOC)) {
            $count_total_rizal1 = $row["counter"];
        }
    }

$count_total_rizal1 = $count_total_rizal1-$count_total_rizal_oos;

$totalVolumeInvoice = $totalVolumeInvoice+$count_total_rizal1;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AC'.$count,$count_total_rizal1)
        -> getStyle('AC'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AD'.$count,$rizalprice)
        -> getStyle('AD'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$amountRIZAL = $rizalprice*$count_total_rizal1;

$totalAmount = $totalAmount+$amountRIZAL;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AE'.$count,'Php'.number_format($amountRIZAL))
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);

//6TH LINE

$count = $count+1;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('A'.$count,strtoupper($sender))
        -> getStyle('A'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('A'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('B'.$count,strtoupper($deltype))
        -> getStyle('B'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('B'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('B'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('C'.$count,'LAGUNA')
        -> getStyle('C'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('C'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('C'.$count)
        -> applyFromArray($border);

$count_total_laguna = 0;

$sql_total_laguna = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_laguna = mysqli_query($db_conn, $sql_total_laguna);
if($query_total_laguna) {
    while ($row = mysqli_fetch_array($query_total_laguna, MYSQLI_ASSOC)) {
            $count_total_laguna = $row["counter"];
        }
    }

$totalVolume = $totalVolume+$count_total_laguna;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('G'.$count,$count_total_laguna)
        -> getStyle('G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G'.$count)
        -> applyFromArray($border);

$count_total_laguna_1day = 0;
$count_total_laguna_2day = 0;
$count_total_laguna_3day = 0;
$count_total_laguna_4day = 0;
$count_total_laguna_5day = 0;
$count_total_laguna_6day = 0;
$count_total_laguna_7day = 0;
$count_total_laguna_8day = 0;
$count_total_laguna_9day = 0;
$count_total_laguna_10day = 0;
$count_total_laguna_11day = 0;
$count_total_laguna_12day = 0;
$count_total_laguna_13day = 0;
$count_total_laguna_14day = 0;
$count_total_laguna_15day = 0;

$sql_total_laguna_1day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -1 DAY) AND record_status_status='1'";
$query_total_laguna_1day = mysqli_query($db_conn, $sql_total_laguna_1day);
if($query_total_laguna_1day) {
    while ($row = mysqli_fetch_array($query_total_laguna_1day, MYSQLI_ASSOC)) {
            $count_total_laguna_1day = $row["counter"];
        }
    }

$sql_total_laguna_2day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -2 DAY) AND record_status_status='1'";
$query_total_laguna_2day = mysqli_query($db_conn, $sql_total_laguna_2day);
if($query_total_laguna_2day) {
    while ($row = mysqli_fetch_array($query_total_laguna_2day, MYSQLI_ASSOC)) {
            $count_total_laguna_2day = $row["counter"];
        }
    }


$sql_total_laguna_3day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -3 DAY) AND record_status_status='1'";
$query_total_laguna_3day = mysqli_query($db_conn, $sql_total_laguna_3day);
if($query_total_laguna_3day) {
    while ($row = mysqli_fetch_array($query_total_laguna_3day, MYSQLI_ASSOC)) {
            $count_total_laguna_3day = $row["counter"];
        }
    }


$sql_total_laguna_4day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -4 DAY) AND record_status_status='1'";
$query_total_laguna_4day = mysqli_query($db_conn, $sql_total_laguna_4day);
if($query_total_laguna_4day) {
    while ($row = mysqli_fetch_array($query_total_laguna_4day, MYSQLI_ASSOC)) {
            $count_total_laguna_4day = $row["counter"];
        }
    }


$sql_total_laguna_5day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -5 DAY) AND record_status_status='1'";
$query_total_laguna_5day = mysqli_query($db_conn, $sql_total_laguna_5day);
if($query_total_laguna_5day) {
    while ($row = mysqli_fetch_array($query_total_laguna_5day, MYSQLI_ASSOC)) {
            $count_total_laguna_5day = $row["counter"];
        }
    }


$sql_total_laguna_6day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -6 DAY) AND record_status_status='1'";
$query_total_laguna_6day = mysqli_query($db_conn, $sql_total_laguna_6day);
if($query_total_laguna_6day) {
    while ($row = mysqli_fetch_array($query_total_laguna_6day, MYSQLI_ASSOC)) {
            $count_total_laguna_6day = $row["counter"];
        }
    }


$sql_total_laguna_7day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -7 DAY) AND record_status_status='1'";
$query_total_laguna_7day = mysqli_query($db_conn, $sql_total_laguna_7day);
if($query_total_laguna_7day) {
    while ($row = mysqli_fetch_array($query_total_laguna_7day, MYSQLI_ASSOC)) {
            $count_total_laguna_7day = $row["counter"];
        }
    }


$sql_total_laguna_8day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -8 DAY) AND record_status_status='1'";
$query_total_laguna_8day = mysqli_query($db_conn, $sql_total_laguna_8day);
if($query_total_laguna_8day) {
    while ($row = mysqli_fetch_array($query_total_laguna_8day, MYSQLI_ASSOC)) {
            $count_total_laguna_8day = $row["counter"];
        }
    }


$sql_total_laguna_9day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -9 DAY) AND record_status_status='1'";
$query_total_laguna_9day = mysqli_query($db_conn, $sql_total_laguna_9day);
if($query_total_laguna_9day) {
    while ($row = mysqli_fetch_array($query_total_laguna_9day, MYSQLI_ASSOC)) {
            $count_total_laguna_9day = $row["counter"];
        }
    }


$sql_total_laguna_10day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -10 DAY) AND record_status_status='1'";
$query_total_laguna_10day = mysqli_query($db_conn, $sql_total_laguna_10day);
if($query_total_laguna_10day) {
    while ($row = mysqli_fetch_array($query_total_laguna_10day, MYSQLI_ASSOC)) {
            $count_total_laguna_10day = $row["counter"];
        }
    }


$sql_total_laguna_11day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -11 DAY) AND record_status_status='1'";
$query_total_laguna_11day = mysqli_query($db_conn, $sql_total_laguna_11day);
if($query_total_laguna_11day) {
    while ($row = mysqli_fetch_array($query_total_laguna_11day, MYSQLI_ASSOC)) {
            $count_total_laguna_11day = $row["counter"];
        }
    }


$sql_total_laguna_12day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -12 DAY) AND record_status_status='1'";
$query_total_laguna_12day = mysqli_query($db_conn, $sql_total_laguna_12day);
if($query_total_laguna_12day) {
    while ($row = mysqli_fetch_array($query_total_laguna_12day, MYSQLI_ASSOC)) {
            $count_total_laguna_12day = $row["counter"];
        }
    }


$sql_total_laguna_13day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -13 DAY) AND record_status_status='1'";
$query_total_laguna_13day = mysqli_query($db_conn, $sql_total_laguna_13day);
if($query_total_laguna_13day) {
    while ($row = mysqli_fetch_array($query_total_laguna_13day, MYSQLI_ASSOC)) {
            $count_total_laguna_13day = $row["counter"];
        }
    }


$sql_total_laguna_14day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -14 DAY) AND record_status_status='1'";
$query_total_laguna_14day = mysqli_query($db_conn, $sql_total_laguna_14day);
if($query_total_laguna_14day) {
    while ($row = mysqli_fetch_array($query_total_laguna_14day, MYSQLI_ASSOC)) {
            $count_total_laguna_14day = $row["counter"];
        }
    }

$sql_total_laguna_15day = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -15 DAY) AND record_status_status='1'";
$query_total_laguna_15day = mysqli_query($db_conn, $sql_total_laguna_15day);
if($query_total_laguna_15day) {
    while ($row = mysqli_fetch_array($query_total_laguna_15day, MYSQLI_ASSOC)) {
            $count_total_laguna_15day = $row["counter"];
        }
    }


$totalDay1 = $totalDay1+$count_total_laguna_1day;
$totalDay2 = $totalDay2+$count_total_laguna_2day;
$totalDay3 = $totalDay3+$count_total_laguna_3day;
$totalDay4 = $totalDay4+$count_total_laguna_4day;
$totalDay5 = $totalDay5+$count_total_laguna_5day;
$totalDay6 = $totalDay6+$count_total_laguna_6day;
$totalDay7 = $totalDay7+$count_total_laguna_7day;
$totalDay8 = $totalDay8+$count_total_laguna_8day;
$totalDay9 = $totalDay9+$count_total_laguna_9day;
$totalDay10 = $totalDay10+$count_total_laguna_10day;
$totalDay11 = $totalDay11+$count_total_laguna_11day;
$totalDay12 = $totalDay12+$count_total_laguna_12day;
$totalDay13 = $totalDay13+$count_total_laguna_13day;
$totalDay14 = $totalDay14+$count_total_laguna_14day;
$totalDay15 = $totalDay15+$count_total_laguna_15day;

$provTotalDay1 = $provTotalDay1+$count_total_laguna_1day;
$provTotalDay2 = $provTotalDay2+$count_total_laguna_2day;
$provTotalDay3 = $provTotalDay3+$count_total_laguna_3day;
$provTotalDay4 = $provTotalDay4+$count_total_laguna_4day;
$provTotalDay5 = $provTotalDay5+$count_total_laguna_5day;
$provTotalDay6 = $provTotalDay6+$count_total_laguna_6day;
$provTotalDay7 = $provTotalDay7+$count_total_laguna_7day;
$provTotalDay8 = $provTotalDay8+$count_total_laguna_8day;
$provTotalDay9 = $provTotalDay9+$count_total_laguna_9day;
$provTotalDay10 = $provTotalDay10+$count_total_laguna_10day;
$provTotalDay11 = $provTotalDay11+$count_total_laguna_11day;
$provTotalDay12 = $provTotalDay12+$count_total_laguna_12day;
$provTotalDay13 = $provTotalDay13+$count_total_laguna_13day;
$provTotalDay14 = $provTotalDay14+$count_total_laguna_14day;
$provTotalDay15 = $provTotalDay15+$count_total_laguna_15day;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('H'.$count,$count_total_laguna_1day)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('I'.$count,$count_total_laguna_2day)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('J'.$count,$count_total_laguna_3day)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('K'.$count,$count_total_laguna_4day)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('L'.$count,$count_total_laguna_5day)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('M'.$count,$count_total_laguna_6day)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('N'.$count,$count_total_laguna_7day)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('O'.$count,$count_total_laguna_8day)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('P'.$count,$count_total_laguna_9day)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Q'.$count,$count_total_laguna_10day)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('R'.$count,$count_total_laguna_11day)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('S'.$count,$count_total_laguna_12day)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('T'.$count,$count_total_laguna_13day)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('U'.$count,$count_total_laguna_14day)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('V'.$count,$count_total_laguna_15day)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$count_total_laguna_delivered = 0;

$sql_total_laguna_delivered = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_status = 'DELIVERED' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_laguna_delivered = mysqli_query($db_conn, $sql_total_laguna_delivered);
if($query_total_laguna_delivered) {
    while ($row = mysqli_fetch_array($query_total_laguna_delivered, MYSQLI_ASSOC)) {
            $count_total_laguna_delivered = $row["counter"];
        }
    }

$totalDelivered = $totalDelivered+$count_total_laguna_delivered;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('W'.$count,$count_total_laguna_delivered)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$count_total_laguna_rts = 0;

$sql_total_laguna_rts = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_status = 'RTS' AND record_reason_rts != 'OUT OF SCOPE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_laguna_rts = mysqli_query($db_conn, $sql_total_laguna_rts);
if($query_total_laguna_rts) {
    while ($row = mysqli_fetch_array($query_total_laguna_rts, MYSQLI_ASSOC)) {
            $count_total_laguna_rts = $row["counter"];
        }
    }

$totalRTS = $totalRTS+$count_total_laguna_rts;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('X'.$count,$count_total_laguna_rts)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$count_total_laguna_oos = 0;

$sql_total_laguna_oos = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_reason_rts = 'OUT OF SCOPE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_laguna_oos = mysqli_query($db_conn, $sql_total_laguna_oos);
if($query_total_laguna_oos) {
    while ($row = mysqli_fetch_array($query_total_laguna_oos, MYSQLI_ASSOC)) {
            $count_total_laguna_oos = $row["counter"];
        }
    }

$totalOOS = $totalOOS+$count_total_laguna_oos;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Y'.$count,$count_total_laguna_oos)
        -> getStyle('Y'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y'.$count)
        -> applyFromArray($border);

$count_total_laguna_balance = 0;

$sql_total_laguna_balance = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_status = '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_laguna_balance = mysqli_query($db_conn, $sql_total_laguna_balance);
if($query_total_laguna_balance) {
    while ($row = mysqli_fetch_array($query_total_laguna_balance, MYSQLI_ASSOC)) {
            $count_total_laguna_balance = $row["counter"];
        }
    }

$totalBalance = $totalBalance+$count_total_laguna_balance;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Z'.$count,$count_total_laguna_balance)
        -> getStyle('Z'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

if($count_total_laguna_balance == 0) {
    $reportStatusLAGUNA = "Complete";
}

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AA'.$count,$reportStatusLAGUNA)
        -> getStyle('AA'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AA'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AA'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AB'.$count,'LAGUNA')
        -> getStyle('AB'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AB'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AB'.$count)
        -> applyFromArray($border);

$count_total_laguna1 = 0;

$sql_total_laguna1 = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_area = 'LAGUNA' AND record_status != '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_laguna1 = mysqli_query($db_conn, $sql_total_laguna1);
if($query_total_laguna1) {
    while ($row = mysqli_fetch_array($query_total_laguna1, MYSQLI_ASSOC)) {
            $count_total_laguna1 = $row["counter"];
        }
    }

$count_total_laguna1 = $count_total_laguna1-$count_total_laguna_oos;

$totalVolumeInvoice = $totalVolumeInvoice+$count_total_laguna1;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AC'.$count,$count_total_laguna1)
        -> getStyle('AC'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AD'.$count,$lagunaprice)
        -> getStyle('AD'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$amountLAGUNA = $lagunaprice*$count_total_laguna1;

$totalAmount = $totalAmount+$amountLAGUNA;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AE'.$count,'Php'.number_format($amountLAGUNA))
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);


$count = $count+1;

}

// FOOTER

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('F'.$count,'TOTAL')
        -> getStyle('F'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('F'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('F'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('G'.$count,$totalVolume)
        -> getStyle('G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('G'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('H'.$count,$totalDay1)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('I'.$count,$totalDay2)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('J'.$count,$totalDay3)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('K'.$count,$totalDay4)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('L'.$count,$totalDay5)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('M'.$count,$totalDay6)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('N'.$count,$totalDay7)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('O'.$count,$totalDay8)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('P'.$count,$totalDay9)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Q'.$count,$totalDay10)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('R'.$count,$totalDay11)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('S'.$count,$totalDay12)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('T'.$count,$totalDay13)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('U'.$count,$totalDay14)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('V'.$count,$totalDay15)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('W'.$count,$totalDelivered)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('X'.$count,$totalRTS)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Y'.$count,$totalOOS)
        -> getStyle('Y'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Z'.$count,$totalBalance)
        -> getStyle('Z'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AA'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('AB'.$count.':AE'.$count);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AB'.$count.':AE'.$count)
        -> applyFromArray($border);

$count = $count+1;

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('F'.$count.':G'.$count);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('F'.$count,'PERCENTAGE')
        -> getStyle('F'.$count.':G'.$count)
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('F'.$count.':G'.$count)
        -> applyFromArray($border);

$day1Percent = get_percentage($totalVolume, $totalDay1);

$day2Percent = get_percentage($totalVolume, $totalDay2);

$day3Percent = get_percentage($totalVolume, $totalDay3);

$day4Percent = get_percentage($totalVolume, $totalDay4);

$day5Percent = get_percentage($totalVolume, $totalDay5);

$day6Percent = get_percentage($totalVolume, $totalDay6);

$day7Percent = get_percentage($totalVolume, $totalDay7);

$day8Percent = get_percentage($totalVolume, $totalDay8);

$day9Percent = get_percentage($totalVolume, $totalDay9);

$day10Percent = get_percentage($totalVolume, $totalDay10);

$day11Percent = get_percentage($totalVolume, $totalDay11);

$day12Percent = get_percentage($totalVolume, $totalDay12);

$day13Percent = get_percentage($totalVolume, $totalDay13);

$day14Percent = get_percentage($totalVolume, $totalDay14);

$day15Percent = get_percentage($totalVolume, $totalDay15);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('H'.$count,$day1Percent.'%')
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('I'.$count,$day2Percent.'%')
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('J'.$count,$day3Percent.'%')
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('K'.$count,$day4Percent.'%')
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('L'.$count,$day5Percent.'%')
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('M'.$count,$day6Percent.'%')
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('N'.$count,$day7Percent.'%')
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('O'.$count,$day8Percent.'%')
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('P'.$count,$day9Percent.'%')
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Q'.$count,$day10Percent.'%')
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('R'.$count,$day11Percent.'%')
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('S'.$count,$day12Percent.'%')
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('T'.$count,$day13Percent.'%')
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('U'.$count,$day14Percent.'%')
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('V'.$count,$day15Percent.'%')
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$deliveredPercent = get_percentage($totalVolume, $totalDelivered);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('W'.$count,$deliveredPercent.'%')
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$rtsPercent = get_percentage($totalVolume, $totalRTS);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('X'.$count,$rtsPercent.'%')
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$oosPercent = get_percentage($totalVolume, $totalOOS);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Y'.$count,$oosPercent.'%')
        -> getStyle('Y'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Y'.$count)
        -> applyFromArray($border);

$balancePercent = get_percentage($totalVolume, $totalBalance);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Z'.$count,$balancePercent.'%')
        -> getStyle('Z'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Z'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AA'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AB'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AC'.$count,$totalVolumeInvoice)
        -> getStyle('AC'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AC'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AE'.$count,'Php'.number_format($totalAmount, 2))
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);

$count = $count+1;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AD'.$count,'ADD 12% Vat')
        -> getStyle('AD'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$vatAmount = $totalAmount*0.12;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AE'.$count,$vatAmount)
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);

$count = $count+1;

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('F'.$count.':V'.$count);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('F'.$count.':V'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('F'.$count.':V'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('366092');

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('W'.$count.':X'.$count);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('W'.$count,'Total')
        -> getStyle('W'.$count.':X'.$count)
        -> applyFromArray($tableHeaderStyleFontYellowBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count.':X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count.':X'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count.':X'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('366092');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AD'.$count,'TOTAL')
        -> getStyle('AD'.$count)
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AD'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$grandTotalAmount = $totalAmount+$vatAmount;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('AE'.$count,$grandTotalAmount)
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('AE'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$count = $count+1;

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('F'.$count.':V'.$count);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('F'.$count.':V'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('F'.$count.':V'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('W'.$count,'On Time')
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('X'.$count,'Late')
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontRedCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$count = $count+1;

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('F'.$count.':G'.$count);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('F'.$count,'Provincial Delivery Status')
        -> getStyle('F'.$count.':G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('F'.$count.':G'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('H'.$count,$provTotalDay1)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('I'.$count,$provTotalDay2)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('J'.$count,$provTotalDay3)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('K'.$count,$provTotalDay4)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('L'.$count,$provTotalDay5)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('M'.$count,$provTotalDay6)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('N'.$count,$provTotalDay7)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('O'.$count,$provTotalDay8)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('P'.$count,$provTotalDay9)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Q'.$count,$provTotalDay10)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('R'.$count,$provTotalDay11)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('S'.$count,$provTotalDay12)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('T'.$count,$provTotalDay13)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('U'.$count,$provTotalDay14)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('V'.$count,$provTotalDay15)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$totalProvOnTime = $provTotalDay1+$provTotalDay2+$provTotalDay3+$provTotalDay4+$provTotalDay5+$provTotalDay6+$provTotalDay7+$provTotalDay8+$provTotalDay9+$provTotalDay10;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('W'.$count,$totalProvOnTime)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$totalProvLate = $provTotalDay11+$provTotalDay12+$provTotalDay13+$provTotalDay14+$provTotalDay15;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('X'.$count,$totalProvLate)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$count = $count+1;

$excel -> setActiveSheetIndex(2) 
        -> mergeCells('F'.$count.':G'.$count);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('F'.$count,'NCR Delivery Status')
        -> getStyle('F'.$count.':G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('F'.$count.':G'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('H'.$count,$ncrTotalDay1)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('H'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('I'.$count,$ncrTotalDay2)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('I'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('J'.$count,$ncrTotalDay3)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('J'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('K'.$count,$ncrTotalDay4)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('K'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('L'.$count,$ncrTotalDay5)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('L'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('M'.$count,$ncrTotalDay6)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('M'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('N'.$count,$ncrTotalDay7)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('N'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('O'.$count,$ncrTotalDay8)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('O'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('P'.$count,$ncrTotalDay9)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('P'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('Q'.$count,$ncrTotalDay10)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('Q'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('R'.$count,$ncrTotalDay11)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('R'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('S'.$count,$ncrTotalDay12)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('S'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('T'.$count,$ncrTotalDay13)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('T'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('U'.$count,$ncrTotalDay14)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('U'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('V'.$count,$ncrTotalDay15)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('V'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$totalNCROnTime = $ncrTotalDay1+$ncrTotalDay2+$ncrTotalDay3+$ncrTotalDay4+$ncrTotalDay5;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('W'.$count,$totalNCROnTime)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$totalNCRLate = $ncrTotalDay6+$ncrTotalDay7+$ncrTotalDay8+$ncrTotalDay9+$ncrTotalDay10+$ncrTotalDay11+$ncrTotalDay12+$ncrTotalDay13+$ncrTotalDay14+$ncrTotalDay15;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('X'.$count,$totalNCRLate)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$totalOnTime = $totalNCROnTime+$totalProvOnTime;

$count = $count+1;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('W'.$count,$totalOnTime)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('W'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('f2f2f2');

$totalLate = $totalNCRLate+$totalProvLate;

$excel -> setActiveSheetIndex(2) 
        -> setCellValue('X'.$count,$totalLate)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri2);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(2) 
        -> getStyle('X'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('f2f2f2');

function get_percentage($total, $number){
  if ( $total > 0 ) {
   return number_format(round($number / ($total / 100),2),2);
  } else {
    return 0;
  }
}

$deltype = preg_replace('/\s+/', '', $deltype);
$sender = preg_replace('/\s+/', '', $sender);

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="'.$sender.''.$deltype.''.$month.''.$year.'Report.xlsx"');
header('Cache-Control: max-age=0');
$file = PHPExcel_IOFactory::createWriter($excel,'Excel2007');
$file -> save('php://output');

mysqli_close($db_conn);

?>