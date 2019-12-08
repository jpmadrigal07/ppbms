<?php

include_once("includes/loginstatus.php");
require_once 'library/PHPExcel/Classes/PHPExcel.php';
header('Content-Type: text/html; charset=ISO-8859-1');
ini_set('max_execution_time', 19200); //300 sedb_connds = 5 minutes
ini_set('memory_limit', '-1');

$excel = new PHPExcel();

// FOURTH SHEET

// Create a new worksheet, after the default sheet
$excel->createSheet();

//SHEET NAME

$excel -> setActiveSheetIndex(3) 
        -> setTitle('Disbursement Summary Report');

//STYLE

$blackCalibriItalic = array(
    'font'  => array(
        'bold'  => false,
        'italic' => true,
        'color' => array('rgb' => '000000'),
        'size'  => 11,
        'name'  => 'Calibri'
    ));

$whiteCalibri = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => 'FFFFFF'),
        'size'  => 11,
        'name'  => 'Calibri'
    ));

$blackCalibri = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => '000000'),
        'size'  => 11,
        'name'  => 'Calibri'
    ));

$blackCalibriThin = array(
    'font'  => array(
        'bold'  => false,
        'color' => array('rgb' => '000000'),
        'size'  => 11,
        'name'  => 'Calibri'
    ));

$redCalibri = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => 'FF0000'),
        'size'  => 11,
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

$borderWhite = array(
        'borders' => array(
          'allborders' => array(
            'style' => PHPExcel_Style_Border::BORDER_THIN,
            'color' => array('rgb' => 'FFFFFF')
          )
        )
      );
      


//TOP HEADER

$excel -> setActiveSheetIndex(3) 
        -> mergeCells('A1:C1');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('A1','Disbursement Summary Report') 
        -> getStyle('A1')
        -> applyFromArray($blackCalibriItalic);

//TABLE HEADER

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('A3','Job Order/COM')
        -> getStyle('A3')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A3')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A3:A5')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A3')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> mergeCells('A3:A5');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('B3','Pick up Date')
        -> getStyle('B3')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('B3')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('B3:B5')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('B3')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> mergeCells('B3:B5');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('C3','Volume Pick up on COM')
        -> getStyle('C3')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C3')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C3:C5')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C3')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> mergeCells('C3:C5');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('D3','Grand Total Delivered')
        -> getStyle('D3')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('D3')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('D3:D5')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('D3')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> mergeCells('D3:D5');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('E3','Total RTS')
        -> getStyle('E3')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('E3')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('E3:E5')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('E3')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> mergeCells('E3:E5');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('F3','Details')
        -> getStyle('F3')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F3')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F3:Q3')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F3')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> mergeCells('F3:Q3');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('F4','NCR')
        -> getStyle('F4')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F4:K4')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F4')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> mergeCells('F4:K4');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('L4','PROVINCIAL')
        -> getStyle('L4')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('L4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('L4:Q4')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('L4')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> mergeCells('L4:Q4');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('F5','Successful')
        -> getStyle('F5')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F5')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F5')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F5')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('G5','RTS')
        -> getStyle('G5')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('G5')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('G5')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('G5')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('H5','OS')
        -> getStyle('H5')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('H5')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('H5')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('H5')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('I5','Total Pick-up')
        -> getStyle('I5')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('I5')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('I5')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('I5')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('J5','Total Delivered')
        -> getStyle('J5')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('J5')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('J5')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('J5')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('K5','Amount/Rate at P 9.00')
        -> getStyle('K5')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('K5')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('K5')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('K5')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

        $excel -> setActiveSheetIndex(3) 
        -> setCellValue('L5','Successful')
        -> getStyle('L5')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('L5')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('L5')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('L5')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('M5','RTS')
        -> getStyle('M5')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('M5')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('M5')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('M5')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('N5','OS')
        -> getStyle('N5')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('N5')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('N5')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('N5')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('O5','Total Pick-up')
        -> getStyle('O5')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('O5')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('O5')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('O5')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('P5','Total Delivered')
        -> getStyle('P5')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('P5')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('P5')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('P5')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('Q5','Amount/Rate at P 10.00')
        -> getStyle('Q5')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('Q5')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('Q5')
        -> applyFromArray($borderWhite);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('Q5')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel->setActiveSheetIndex(3)->getColumnDimension('A')->setWidth(20);
$excel->setActiveSheetIndex(3)->getColumnDimension('B')->setWidth(12);
$excel->setActiveSheetIndex(3)->getColumnDimension('C')->setWidth(23);
$excel->setActiveSheetIndex(3)->getColumnDimension('D')->setWidth(23);
$excel->setActiveSheetIndex(3)->getColumnDimension('E')->setWidth(10);
$excel->setActiveSheetIndex(3)->getColumnDimension('F')->setWidth(10);
$excel->setActiveSheetIndex(3)->getColumnDimension('G')->setWidth(8);
$excel->setActiveSheetIndex(3)->getColumnDimension('H')->setWidth(6);
$excel->setActiveSheetIndex(3)->getColumnDimension('I')->setWidth(13);
$excel->setActiveSheetIndex(3)->getColumnDimension('J')->setWidth(15);
$excel->setActiveSheetIndex(3)->getColumnDimension('K')->setWidth(23);
$excel->setActiveSheetIndex(3)->getColumnDimension('L')->setWidth(10);
$excel->setActiveSheetIndex(3)->getColumnDimension('M')->setWidth(8);
$excel->setActiveSheetIndex(3)->getColumnDimension('N')->setWidth(6);
$excel->setActiveSheetIndex(3)->getColumnDimension('O')->setWidth(13);
$excel->setActiveSheetIndex(3)->getColumnDimension('P')->setWidth(15);
$excel->setActiveSheetIndex(3)->getColumnDimension('Q')->setWidth(23);

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('A6','20190830')
        -> getStyle('A6')
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('B6','9/2/2019')
        -> getStyle('B6')
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('B6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('B6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('B6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('C6','0')
        -> getStyle('C6')
        -> applyFromArray($blackCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('FFFFFF');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('D6','0')
        -> getStyle('D6')
        -> applyFromArray($blackCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('D6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('D6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('D6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('FFFFFF');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('E6','0')
        -> getStyle('E6')
        -> applyFromArray($blackCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('E6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('E6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('E6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('FFFFFF');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('F6','0')
        -> getStyle('F6')
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('G6','0')
        -> getStyle('G6')
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('G6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('G6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('G6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('H6','0')
        -> getStyle('H6')
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('H6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('H6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('H6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('I6','0')
        -> getStyle('I6')
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('I6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('I6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('I6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('FFFFFF');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('J6','0')
        -> getStyle('J6')
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('J6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('J6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('J6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('FFFFFF');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('K6','0')
        -> getStyle('K6')
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('K6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('K6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('K6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('e4dfec');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('L6','0')
        -> getStyle('L6')
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('L6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('L6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('L6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('M6','0')
        -> getStyle('M6')
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('M6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('M6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('M6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('N6','0')
        -> getStyle('N6')
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('N6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('N6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('N6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('O6','0')
        -> getStyle('O6')
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('O6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('O6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('O6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('FFFFFF');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('P6','0')
        -> getStyle('P6')
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('P6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('P6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('P6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('FFFFFF');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('Q6','0')
        -> getStyle('Q6')
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('Q6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('Q6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('Q6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('e4dfec');

// FOOTER

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('A8','NCR')
        -> getStyle('A8')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('B8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('C8','0')
        -> getStyle('C8')
        -> applyFromArray($blackCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('FFFFFF');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C8:Q8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('A9','PROVINCIAL')
        -> getStyle('A9')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A9')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A9')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A9')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('B9')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('C9','0')
        -> getStyle('C9')
        -> applyFromArray($blackCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C9')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C9')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('FFFFFF');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C9:Q9')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('A10','TOTAL')
        -> getStyle('A10')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A10')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A10')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A10')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('B10')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('C10','0')
        -> getStyle('C10')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C10')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C10')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('D10','0')
        -> getStyle('D10')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('D10')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('D10')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('E10','0')
        -> getStyle('E10')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('E10')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('E10')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('F10','0')
        -> getStyle('F10')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F10')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F10')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('G10','0')
        -> getStyle('G10')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('G10')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('G10')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('H10','0')
        -> getStyle('H10')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('H10')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('H10')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('I10')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('I10')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('J10','0')
        -> getStyle('J10')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('J10')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('J10')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('K10','0')
        -> getStyle('K10')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('K10')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('K10')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('L10','0')
        -> getStyle('L10')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('L10')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('L10')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('M10','0')
        -> getStyle('M10')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('M10')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('M10')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('N10','0')
        -> getStyle('N10')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('N10')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('N10')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('O10')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('O10')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('P10','0')
        -> getStyle('P10')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('P10')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('P10')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('Q10','0')
        -> getStyle('Q10')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('Q10')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('Q10')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('K11','0')
        -> getStyle('K11')
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('K11')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('Q11','0')
        -> getStyle('Q11')
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('Q11')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('A12','GRAND TOTAL')
        -> getStyle('A12')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A12')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A12')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('B12')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C12')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('D12')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('E12')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F12')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('G12')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('H12')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('I12')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('J12')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('K12','0')
        -> getStyle('K12')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('K12')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('K12')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('L12')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('M12')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('N12')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('O12')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('P12')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('Q12','0')
        -> getStyle('Q12')
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('Q12')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('Q12')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('R12','0')
        -> getStyle('R12')
        -> applyFromArray($redCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('R12')
        -> applyFromArray($centerText);


// DATA LOOP

$countStart = 3;
$footerCount = $countSummary+1;
$totalDeliveredFinal = 0;
$totalRTSFinal = 0;
$totalDeliveredNCR = 0;
$totalRTSNCR = 0;
$totalOOSNCR = 0;
$totalPickupNCR = 0;
$totalPickupNCRFinal = 0;
$totalSuccessfulRTSNCR = 0;
$totalSuccessfulRTSNCRFinal = 0;
$totalSuccessfulRTSNCRPayment = 0;
$totalSuccessfulRTSNCRPaymentFinal = 0;
$totalDeliveredProv = 0;
$totalRTSProv = 0;
$totalOOSProv = 0;
$totalPickupProv = 0;
$totalPickupProvFinal = 0;
$totalSuccessfulRTSProv = 0;
$totalSuccessfulRTSProvFinal = 0;
$totalSuccessfulRTSProvPayment = 0;
$totalSuccessfulRTSProvPaymentFinal = 0;


for($i1=0; $i1<$countSummary; $i1++) {
    $summaryArray = explode(',', $_POST['checkboxsummary'][$i1]);
    $cyclecode = $summaryArray[0];
    $pud = $summaryArray[1];
    $countStart++;

    // COUNT PROVINCIAL VOLUME END

    $excel -> setActiveSheetIndex(3) 
        -> setCellValue('A'.$countStart,$cyclecode)
        -> getStyle('A'.$countStart)
        -> applyFromArray($blackCalibriThin);

    $excel -> setActiveSheetIndex(3) 
        -> getStyle('A'.$countStart)
        -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(3) 
        -> getStyle('A'.$countStart)
        -> applyFromArray($border);

    $excel -> setActiveSheetIndex(3) 
        -> getStyle('A'.$countStart)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

    // DIVIDER

    $excel -> setActiveSheetIndex(3) 
        -> setCellValue('B'.$countStart,$pud)
        -> getStyle('B'.$countStart)
        -> applyFromArray($blackCalibriThin);

    $excel -> setActiveSheetIndex(3) 
        -> getStyle('B'.$countStart)
        -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(3) 
        -> getStyle('B'.$countStart)
        -> applyFromArray($border);

    $excel -> setActiveSheetIndex(3) 
        -> getStyle('B'.$countStart)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

    // DIVIDER

    $count_volume = 0;

    $sql_volume = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_volume = mysqli_query($db_conn, $sql_volume);
    if($query_volume) {
        while ($row = mysqli_fetch_array($query_volume, MYSQLI_ASSOC)) {
            $count_volume = $row["counter"];
        }
    }

    $excel -> setActiveSheetIndex(3) 
        -> setCellValue('C'.$countStart,$count_volume)
        -> getStyle('C'.$countStart)
        -> applyFromArray($blackCalibri);

    $excel -> setActiveSheetIndex(3) 
        -> getStyle('C'.$countStart)
        -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(3) 
        -> getStyle('C'.$countStart)
        -> applyFromArray($border);
   
    $excel -> setActiveSheetIndex(3) 
        -> getStyle('C'.$countStart)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('FFFFFF');

    // DIVIDER

    $count_delivered = 0;

    $sql_delivered = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_status = 'DELIVERED' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_delivered = mysqli_query($db_conn, $sql_delivered);
    if($query_delivered) {
        while ($row = mysqli_fetch_array($query_delivered, MYSQLI_ASSOC)) {
            $count_delivered = $row["counter"];
        }
    }

    $totalDeliveredFinal = $totalDeliveredFinal+$count_delivered;

    $excel -> setActiveSheetIndex(3) 
            -> setCellValue('D'.$countStart,$count_delivered)
            -> getStyle('D'.$countStart)
            -> applyFromArray($blackCalibri);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('D'.$countStart)
            -> applyFromArray($centerText);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('D'.$countStart)
            -> applyFromArray($border);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('D'.$countStart)
            -> getFill()
            -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            -> getStartColor()
            -> setRGB('FFFFFF');

    // DIVIDER

    $count_rts = 0;

    $sql_rts = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_status = 'RTS' AND record_reason_rts NOT LIKE '%OUT OF SCOPE%' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_rts = mysqli_query($db_conn, $sql_rts);
    if($query_rts) {
        while ($row = mysqli_fetch_array($query_rts, MYSQLI_ASSOC)) {
            $count_rts = $row["counter"];
        }
    }

    $totalRTSFinal = $totalRTSFinal+$count_rts;

    $excel -> setActiveSheetIndex(3) 
        -> setCellValue('E'.$countStart,$count_rts)
        -> getStyle('E'.$countStart)
        -> applyFromArray($topHeaderStyleFontBlackCalibri1);

    $excel -> setActiveSheetIndex(3) 
            -> getStyle('E'.$countStart)
            -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(3) 
            -> getStyle('E'.$countStart)
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

    $totalDeliveredNCR = $totalDeliveredNCR+$count_delivered_ncr;

    $excel -> setActiveSheetIndex(3) 
            -> setCellValue('F'.$countStart,$count_delivered_ncr)
            -> getStyle('F'.$countStart)
            -> applyFromArray($blackCalibriThin);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('F'.$countStart)
            -> applyFromArray($centerText);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('F'.$countStart)
            -> applyFromArray($border);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('F'.$countStart)
            -> getFill()
            -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            -> getStartColor()
            -> setRGB('ddd9c4');

    // DIVIDER

    $count_rts_ncr = 0;

    $sql_rts_ncr = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_status = 'RTS' AND record_reason_rts NOT LIKE '%OUT OF SCOPE%' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE' OR record_area = 'MARIKINA' OR record_area = 'MUNTINLUPA' OR record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_rts_ncr = mysqli_query($db_conn, $sql_rts_ncr);
    if($query_rts_ncr) {
        while ($row = mysqli_fetch_array($query_rts_ncr, MYSQLI_ASSOC)) {
            $count_rts_ncr = $row["counter"];
        }
    }

    $totalRTSNCR = $totalRTSNCR+$count_rts_ncr;

    $excel -> setActiveSheetIndex(3) 
            -> setCellValue('G'.$countStart,$count_rts_ncr)
            -> getStyle('G'.$countStart)
            -> applyFromArray($blackCalibriThin);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('G'.$countStart)
            -> applyFromArray($centerText);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('G'.$countStart)
            -> applyFromArray($border);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('G'.$countStart)
            -> getFill()
            -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            -> getStartColor()
            -> setRGB('ddd9c4');

    // DIVIDER

    $count_oos_ncr = 0;

    $sql_oos_ncr = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_status = 'RTS' AND record_reason_rts LIKE '%OUT OF SCOPE%' AND (record_area = 'PARANAQUE' OR record_area = 'PARA?AQUE' OR record_area = 'PARAÑAQUE' OR record_area = 'MARIKINA' OR record_area = 'MUNTINLUPA' OR record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_oos_ncr = mysqli_query($db_conn, $sql_oos_ncr);
    if($query_oos_ncr) {
        while ($row = mysqli_fetch_array($query_oos_ncr, MYSQLI_ASSOC)) {
            $count_oos_ncr = $row["counter"];
        }
    }

    $totalOOSNCR = $totalOOSNCR+$count_oos_ncr;

    $excel -> setActiveSheetIndex(3) 
            -> setCellValue('H'.$countStart,$count_oos_ncr)
            -> getStyle('H'.$countStart)
            -> applyFromArray($blackCalibriThin);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('H'.$countStart)
            -> applyFromArray($centerText);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('H'.$countStart)
            -> applyFromArray($border);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('H'.$countStart)
            -> getFill()
            -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            -> getStartColor()
            -> setRGB('ddd9c4');

    $totalPickupNCR = $count_delivered_ncr+$count_rts_ncr+$count_oos_ncr;
    $totalPickupNCRFinal = $totalPickupNCRFinal+$totalPickupNCR;

    $excel -> setActiveSheetIndex(3) 
            -> setCellValue('I'.$countStart,$totalPickupNCR)
            -> getStyle('I'.$countStart)
            -> applyFromArray($blackCalibriThin);

    $excel -> setActiveSheetIndex(3) 
            -> getStyle('I'.$countStart)
            -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(3) 
            -> getStyle('I'.$countStart)
            -> applyFromArray($border);

    $excel -> setActiveSheetIndex(3) 
            -> getStyle('I'.$countStart)
            -> getFill()
            -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            -> getStartColor()
            -> setRGB('FFFFFF');

    $totalSuccessfulRTSNCR = $count_delivered_ncr+$count_rts_ncr;
    $totalSuccessfulRTSNCRFinal = $totalSuccessfulRTSNCRFinal+$totalSuccessfulRTSNCR;

    $excel -> setActiveSheetIndex(3) 
            -> setCellValue('J'.$countStart,$totalSuccessfulRTSNCR)
            -> getStyle('J'.$countStart)
            -> applyFromArray($blackCalibriThin);

    $excel -> setActiveSheetIndex(3) 
            -> getStyle('J'.$countStart)
            -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(3) 
            -> getStyle('J'.$countStart)
            -> applyFromArray($border);

    $excel -> setActiveSheetIndex(3) 
            -> getStyle('J'.$countStart)
            -> getFill()
            -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            -> getStartColor()
            -> setRGB('FFFFFF');

    $totalSuccessfulRTSNCRPayment = $totalSuccessfulRTSNCR*$pqueprice;
    $totalSuccessfulRTSNCRPaymentFinal = $totalSuccessfulRTSNCRPaymentFinal+$totalSuccessfulRTSNCRPayment;

    $excel -> setActiveSheetIndex(3) 
            -> setCellValue('K'.$countStart,number_format($totalSuccessfulRTSNCRPayment))
            -> getStyle('K'.$countStart)
            -> applyFromArray($blackCalibriThin);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('K'.$countStart)
            -> applyFromArray($centerText);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('K'.$countStart)
            -> applyFromArray($border);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('K'.$countStart)
            -> getFill()
            -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            -> getStartColor()
            -> setRGB('e4dfec');

    // DIVIDER

    $count_delivered_prov = 0;

    $sql_delivered_prov = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_status = 'DELIVERED' AND (record_area = 'LAGUNA' OR record_area = 'RIZAL') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_delivered_prov = mysqli_query($db_conn, $sql_delivered_prov);
    if($query_delivered_prov) {
        while ($row = mysqli_fetch_array($query_delivered_prov, MYSQLI_ASSOC)) {
            $count_delivered_prov = $row["counter"];
        }
    }

    $totalDeliveredProv = $totalDeliveredProv+$count_delivered_prov;

    $excel -> setActiveSheetIndex(3) 
            -> setCellValue('L'.$countStart,$count_delivered_prov)
            -> getStyle('L'.$countStart)
            -> applyFromArray($blackCalibriThin);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('L'.$countStart)
            -> applyFromArray($centerText);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('L'.$countStart)
            -> applyFromArray($border);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('L'.$countStart)
            -> getFill()
            -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            -> getStartColor()
            -> setRGB('ddd9c4');

    // DIVIDER

    $count_rts_prov = 0;

    $sql_rts_prov = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_status = 'RTS' AND record_reason_rts NOT LIKE '%OUT OF SCOPE%' AND (record_area = 'LAGUNA' OR record_area = 'RIZAL') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_rts_prov = mysqli_query($db_conn, $sql_rts_prov);
    if($query_rts_prov) {
        while ($row = mysqli_fetch_array($query_rts_prov, MYSQLI_ASSOC)) {
            $count_rts_prov = $row["counter"];
        }
    }

    $totalRTSProv = $totalRTSProv+$count_rts_prov;

    $excel -> setActiveSheetIndex(3) 
            -> setCellValue('M'.$countStart,$count_rts_prov)
            -> getStyle('M'.$countStart)
            -> applyFromArray($blackCalibriThin);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('M'.$countStart)
            -> applyFromArray($centerText);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('M'.$countStart)
            -> applyFromArray($border);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('M'.$countStart)
            -> getFill()
            -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            -> getStartColor()
            -> setRGB('ddd9c4');

    // DIVIDER

    $count_oos_prov = 0;

    $sql_oos_prov = "SELECT COUNT(id) AS counter FROM ppbms_record WHERE record_sender='$sender' AND record_deltype='$deltype' AND record_month='$month' AND record_year='$year' AND record_status = 'RTS' AND record_reason_rts LIKE '%OUT OF SCOPE%' AND (record_area = 'LAGUNA' OR record_area = 'RIZAL') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
    $query_oos_prov = mysqli_query($db_conn, $sql_oos_prov);
    if($query_oos_prov) {
        while ($row = mysqli_fetch_array($query_oos_prov, MYSQLI_ASSOC)) {
            $count_oos_prov = $row["counter"];
        }
    }

    $totalOOSProv = $totalOOSProv+$count_oos_prov;

    $excel -> setActiveSheetIndex(3) 
            -> setCellValue('N'.$countStart,$count_oos_prov)
            -> getStyle('N'.$countStart)
            -> applyFromArray($blackCalibriThin);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('N'.$countStart)
            -> applyFromArray($centerText);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('N'.$countStart)
            -> applyFromArray($border);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('N'.$countStart)
            -> getFill()
            -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            -> getStartColor()
            -> setRGB('ddd9c4');

    $totalPickupProv = $count_delivered_prov+$count_rts_prov+$count_oos_prov;
    $totalPickupProvFinal = $totalPickupProvFinal+$totalPickupProv;

    $excel -> setActiveSheetIndex(3) 
            -> setCellValue('O'.$countStart,$totalPickupProv)
            -> getStyle('O'.$countStart)
            -> applyFromArray($blackCalibriThin);

    $excel -> setActiveSheetIndex(3) 
            -> getStyle('O'.$countStart)
            -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(3) 
            -> getStyle('O'.$countStart)
            -> applyFromArray($border);

    $excel -> setActiveSheetIndex(3) 
            -> getStyle('O'.$countStart)
            -> getFill()
            -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            -> getStartColor()
            -> setRGB('FFFFFF');

    $totalSuccessfulRTSProv = $count_delivered_prov+$count_rts_prov;
    $totalSuccessfulRTSProvFinal = $totalSuccessfulRTSProvFinal+$totalSuccessfulRTSProv;

    $excel -> setActiveSheetIndex(3) 
            -> setCellValue('P'.$countStart,$totalSuccessfulRTSProv)
            -> getStyle('P'.$countStart)
            -> applyFromArray($blackCalibriThin);

    $excel -> setActiveSheetIndex(3) 
            -> getStyle('P'.$countStart)
            -> applyFromArray($centerText);

    $excel -> setActiveSheetIndex(3) 
            -> getStyle('P'.$countStart)
            -> applyFromArray($border);

    $excel -> setActiveSheetIndex(3) 
            -> getStyle('P'.$countStart)
            -> getFill()
            -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            -> getStartColor()
            -> setRGB('FFFFFF');

    $totalSuccessfulRTSProvPayment = $totalSuccessfulRTSProv*$rizalprice;
    $totalSuccessfulRTSProvPaymentFinal = $totalSuccessfulRTSProvPaymentFinal+$totalSuccessfulRTSProvPayment;

    $excel -> setActiveSheetIndex(3) 
            -> setCellValue('Q'.$countStart,number_format($totalSuccessfulRTSProvPayment))
            -> getStyle('Q'.$countStart)
            -> applyFromArray($blackCalibriThin);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('Q'.$countStart)
            -> applyFromArray($centerText);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('Q'.$countStart)
            -> applyFromArray($border);
    
    $excel -> setActiveSheetIndex(3) 
            -> getStyle('Q'.$countStart)
            -> getFill()
            -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            -> getStartColor()
            -> setRGB('e4dfec');

}

// FOOTER

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('A'.$footerCount,'NCR')
        -> getStyle('A'.$footerCount)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A'.$footerCount)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A'.$footerCount)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A'.$footerCount)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('B'.$footerCount)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('C'.$footerCount,$totalPickupNCRFinal)
        -> getStyle('C'.$footerCount)
        -> applyFromArray($blackCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C'.$footerCount)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C'.$footerCount)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('FFFFFF');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C'.$footerCount.':Q'.$footerCount)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('A'.$footerCount+1,'PROVINCIAL')
        -> getStyle('A'.$footerCount+1)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A'.$footerCount+1)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A'.$footerCount+1)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A'.$footerCount+1)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('B'.$footerCount+1)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('C'.$footerCount+1,$totalPickupProvFinal)
        -> getStyle('C'.$footerCount+1)
        -> applyFromArray($blackCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C'.$footerCount+1)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C'.$footerCount+1)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('FFFFFF');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C'.$footerCount+1.':Q'.$footerCount+1)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('A'.$footerCount+2,'TOTAL')
        -> getStyle('A'.$footerCount+2)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A'.$footerCount+2)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A'.$footerCount+2)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A'.$footerCount+2)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('B'.$footerCount+2)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('C'.$footerCount+2,$totalPickupNCRFinal+$totalPickupProvFinal)
        -> getStyle('C'.$footerCount+2)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C'.$footerCount+2)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C'.$footerCount+2)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('D'.$footerCount+2,$totalDeliveredFinal)
        -> getStyle('D'.$footerCount+2)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('D'.$footerCount+2)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('D'.$footerCount+2)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('E'.$footerCount+2,$totalRTSFinal)
        -> getStyle('E'.$footerCount+2)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('E'.$footerCount+2)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('E'.$footerCount+2)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('F'.$footerCount+2,$totalDeliveredNCR)
        -> getStyle('F'.$footerCount+2)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F'.$footerCount+2)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F'.$footerCount+2)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('G'.$footerCount+2,$totalRTSNCR)
        -> getStyle('G'.$footerCount+2)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('G'.$footerCount+2)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('G'.$footerCount+2)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('H'.$footerCount+2,$totalOOSNCR)
        -> getStyle('H'.$footerCount+2)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('H'.$footerCount+2)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('H'.$footerCount+2)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('I'.$footerCount+2)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('J'.$footerCount+2,$totalSuccessfulRTSNCRFinal)
        -> getStyle('J'.$footerCount+2)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('J'.$footerCount+2)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('J'.$footerCount+2)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('K'.$footerCount+2,number_format($totalSuccessfulRTSNCRPaymentFinal))
        -> getStyle('K'.$footerCount+2)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('K'.$footerCount+2)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('K'.$footerCount+2)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('L'.$footerCount+2,$totalDeliveredProv)
        -> getStyle('L'.$footerCount+2)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('L'.$footerCount+2)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('L'.$footerCount+2)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('M'.$footerCount+2,$totalRTSProv)
        -> getStyle('M'.$footerCount+2)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('M'.$footerCount+2)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('M'.$footerCount+2)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('N'.$footerCount+2,$totalOOSProv)
        -> getStyle('N'.$footerCount+2)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('N'.$footerCount+2)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('N'.$footerCount+2)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('O'.$footerCount+2)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('P'.$footerCount+2,$totalSuccessfulRTSProvFinal)
        -> getStyle('P'.$footerCount+2)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('P'.$footerCount+2)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('P'.$footerCount+2)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('Q'.$footerCount+2,number_format($totalSuccessfulRTSProvPaymentFinal))
        -> getStyle('Q'.$footerCount+2)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('Q'.$footerCount+2)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('Q'.$footerCount+2)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');



$excel -> setActiveSheetIndex(3) 
        -> setCellValue('K'.$footerCount+3, number_format($totalSuccessfulRTSNCRPaymentFinal*0.12))
        -> getStyle('K'.$footerCount+3)
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('K'.$footerCount+3)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('Q'.$footerCount+3, number_format($totalSuccessfulRTSProvPaymentFinal*0.12))
        -> getStyle('Q'.$footerCount+3)
        -> applyFromArray($blackCalibriThin);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('Q'.$footerCount+3)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('A'.$footerCount+4,'GRAND TOTAL')
        -> getStyle('A'.$footerCount+4)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A'.$footerCount+4)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('A'.$footerCount+4)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('B'.$footerCount+4)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('C'.$footerCount+4)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('D'.$footerCount+4)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('E'.$footerCount+4)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('F'.$footerCount+4)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('G'.$footerCount+4)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('H'.$footerCount+4)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('I'.$footerCount+4)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('J'.$footerCount+4)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$grandTotalNCR = $totalSuccessfulRTSNCRPaymentFinal+($totalSuccessfulRTSNCRPaymentFinal*0.12);

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('K'.$footerCount+4,number_format($grandTotalNCR))
        -> getStyle('K'.$footerCount+4)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('K'.$footerCount+4)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('K'.$footerCount+4)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('L'.$footerCount+4)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('M'.$footerCount+4)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('N'.$footerCount+4)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('O'.$footerCount+4)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> getStyle('P'.$footerCount+4)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$grandTotalProv = $totalSuccessfulRTSProvPaymentFinal+($totalSuccessfulRTSProvPaymentFinal*0.12);

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('Q'.$footerCount+4,number_format($grandTotalProv))
        -> getStyle('Q'.$footerCount+4)
        -> applyFromArray($whiteCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('Q'.$footerCount+4)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('Q'.$footerCount+4)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel -> setActiveSheetIndex(3) 
        -> setCellValue('R'.$footerCount+4, number_format($grandTotalProv+$grandTotalNCR))
        -> getStyle('R'.$footerCount+4)
        -> applyFromArray($redCalibri);

$excel -> setActiveSheetIndex(3) 
        -> getStyle('R'.$footerCount+4)
        -> applyFromArray($centerText);

        
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="Report.xlsx"');
header('Cache-Control: max-age=0');
$file = PHPExcel_IOFactory::createWriter($excel,'Excel2007');
$file -> save('php://output');

?>