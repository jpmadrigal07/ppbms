<?php
include_once("includes/loginstatus.php");
require_once 'library/PHPExcel/Classes/PHPExcel.php';
$excel = new PHPExcel();

//SHEET NAME

$excel -> setActiveSheetIndex(0) 
        -> setTitle('Days of Delivery');

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

$tableHeaderStyleFontBlackBoldCalibri = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => '000000'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$tableHeaderStyleFontBlackCalibri = array(
    'font'  => array(
        'bold'  => false,
        'color' => array('rgb' => '000000'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$tableHeaderStyleFontRedBoldCalibri = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => 'ff0000'),
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

$tableHeaderStyleFontYellowBoldCalibri = array(
    'font'  => array(
        'bold'  => true,
        'color' => array('rgb' => 'ffff00'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$tableHeaderStyleFontYellowCalibri = array(
    'font'  => array(
        'bold'  => false,
        'color' => array('rgb' => 'ffff00'),
        'size'  => 10,
        'name'  => 'Calibri'
    ));

$centerText = array(
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER_CONTINUOUS
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

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('A1:Z1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('A1',strtoupper($month).' '.strtoupper($sender).'(OFFICIAL RECEIPT) PERFORMANCE EVALUATION') 
        -> getStyle('A1')
        -> applyFromArray($topHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A1')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0)
        -> getStyle('A1')->getFont()->setSize(24);

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('A3:Z3');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('A3','F L O W  OF  D E L I V E R Y') 
        -> getStyle('A3')
        -> applyFromArray($topHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A3')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0)
        -> getStyle('A3')->getFont()->setSize(16);

//TABLE HEADER

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('A4:A8');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('A4','SENDER')
        -> getStyle('A4:A8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A4:A8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A4:A8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A4:A8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('B4:B8');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('B4','DELIVERY TYPE')
        -> getStyle('B4:B8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B4:B8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B4:B8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B4:B8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('C4:C8');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('C4','AREA')
        -> getStyle('C4:C8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C4:C8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C4:C8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C4:C8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('D4:D8');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('D4','MONTH')
        -> getStyle('D4:D8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('D4:D8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('D4:D8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('D4:D8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('E4:E8');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('E4','CYCLE CODE')
        -> getStyle('E4:E8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('E4:E8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('E4:E8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('E4:E8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('F4:F8');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('F4','PICK-UP DATE')
        -> getStyle('F4:F8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F4:F8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F4:F8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F4:F8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F4:F8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('G4:G8');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('G4','VOLUME')
        -> getStyle('G4:G8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G4:G8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G4:G8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G4:G8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G4:G8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('H4:V4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('H4','DELIVERY TIMELINE')
        -> getStyle('H4:V4')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H4:V4')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H4:V4')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H4:V4')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H4:V4')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('H5:Q5');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('H5','ON TIME DELIVERY FOR PROVINCIAL')
        -> getStyle('H5:Q5')
        -> applyFromArray($tableHeaderStyleFontYellowBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H5:Q5')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H5:Q5')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H5:Q5')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H5:Q5')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('0070c0');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('R5:V5');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('R5','LATE DELIVERY FOR PROVINCIAL')
        -> getStyle('R5:V5')
        -> applyFromArray($tableHeaderStyleFontYellowBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R5:V5')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R5:V5')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R5:V5')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R5:V5')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('0070c0');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('H6','DAY 1')
        -> getStyle('H6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('I6','DAY 2')
        -> getStyle('I6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('J6','DAY 3')
        -> getStyle('J6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('K6','DAY 4')
        -> getStyle('K6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('L6','DAY 5')
        -> getStyle('L6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('M6','DAY 6')
        -> getStyle('M6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('N6','DAY 7')
        -> getStyle('N6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('O6','DAY 8')
        -> getStyle('O6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('P6','DAY 9')
        -> getStyle('P6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Q6','DAY 10')
        -> getStyle('Q6')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('R6','DAY 11')
        -> getStyle('R6')
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('S6','DAY 12')
        -> getStyle('S6')
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('T6','DAY 13')
        -> getStyle('T6')
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('U6','DAY 14')
        -> getStyle('U6')
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('V6','DAY 15')
        -> getStyle('V6')
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V6')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V6')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V6')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V6')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('H7:L7');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('H7','ON TIME DELIVERY FOR NCR')
        -> getStyle('H7:L7')
        -> applyFromArray($tableHeaderStyleFontYellowBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H7:L7')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H7:L7')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H7:L7')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H7:L7')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('0070c0');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('M7:V7');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('M7','LATE DELIVERY FOR NCR')
        -> getStyle('M7:V7')
        -> applyFromArray($tableHeaderStyleFontYellowBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M7:V7')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M7:V7')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M7:V7')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M7:V7')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('0070c0');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('H8','DAY 1')
        -> getStyle('H8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('I8','DAY 2')
        -> getStyle('I8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('J8','DAY 3')
        -> getStyle('J8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('K8','DAY 4')
        -> getStyle('K8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('L8','DAY 5')
        -> getStyle('L8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('M8','DAY 6')
        -> getStyle('M8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('N8','DAY 7')
        -> getStyle('N8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('O8','DAY 8')
        -> getStyle('O8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('P8','DAY 9')
        -> getStyle('P8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Q8','DAY 10')
        -> getStyle('Q8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('R8','DAY 11')
        -> getStyle('R8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('S8','DAY 12')
        -> getStyle('S8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('T8','DAY 13')
        -> getStyle('T8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('U8','DAY 14')
        -> getStyle('U8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('V8','DAY 15')
        -> getStyle('V8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ddd9c4');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('W4:W8');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('W4',"DELIVERED")
        -> getStyle('W4:W8')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W4:W8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W4:W8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W4:W8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W4:W8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('X4:X8');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('X4','RTS')
        -> getStyle('X4:X8')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X4:X8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X4:X8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X4:X8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('Y4:Y8');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Y4','OUT OF SCOPE')
        -> getStyle('Y4:Y8')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y4:Y8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y4:Y8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y4:Y8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('Z4:Z8');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Z4','BALANCE')
        -> getStyle('Z4:Z8')
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z4:Z8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z4:Z8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z4:Z8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z4:Z8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('AA4:AA8');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AA4',"REPORT STATUS")
        -> getStyle('AA4:AA8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AA4:AA8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AA4:AA8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AA4:AA8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('AB4:AB8');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AB4','AREA')
        -> getStyle('AB4:AB8')
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AB4:AB8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AB4:AB8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AB4:AB8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('AC4:AC8');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AC4',"TOTAL VOLUME INVOICE")
        -> getStyle('AC4:AC8')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC4:AC8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC4:AC8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC4:AC8')
        -> getAlignment()
        -> setWrapText(true);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC4:AC8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('AD4:AD8');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AD4','PRICE')
        -> getStyle('AD4:AD8')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD4:AD8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD4:AD8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD4:AD8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('AE4:AE8');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AE4','AMOUNT')
        -> getStyle('AE4:AE8')
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE4:AE8')
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE4:AE8')
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE4:AE8')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('00b050');


$excel -> setActiveSheetIndex(0) 
        -> mergeCells('A9:AE9');

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A9:AE9')
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$excel->setActiveSheetIndex(0)->getRowDimension('4')->setRowHeight(-1);
$excel->setActiveSheetIndex(0)->getRowDimension('5')->setRowHeight(-1);
$excel->setActiveSheetIndex(0)->getRowDimension('6')->setRowHeight(-1);
$excel->setActiveSheetIndex(0)->getRowDimension('7')->setRowHeight(-1);
$excel->setActiveSheetIndex(0)->getRowDimension('8')->setRowHeight(-1);
$excel->setActiveSheetIndex(0)->getColumnDimension('A')->setWidth(9);
$excel->setActiveSheetIndex(0)->getColumnDimension('B')->setWidth(15);
$excel->setActiveSheetIndex(0)->getColumnDimension('C')->setWidth(8);
$excel->setActiveSheetIndex(0)->getColumnDimension('D')->setWidth(10);
$excel->setActiveSheetIndex(0)->getColumnDimension('E')->setWidth(13);
$excel->setActiveSheetIndex(0)->getColumnDimension('F')->setWidth(13);
$excel->setActiveSheetIndex(0)->getColumnDimension('G')->setWidth(9);
$excel->setActiveSheetIndex(0)->getColumnDimension('W')->setWidth(11);
$excel->setActiveSheetIndex(0)->getColumnDimension('X')->setWidth(9);
$excel->setActiveSheetIndex(0)->getColumnDimension('Y')->setWidth(13);
$excel->setActiveSheetIndex(0)->getColumnDimension('Z')->setWidth(11);
$excel->setActiveSheetIndex(0)->getColumnDimension('AA')->setWidth(14);
$excel->setActiveSheetIndex(0)->getColumnDimension('AB')->setWidth(12);
$excel->setActiveSheetIndex(0)->getColumnDimension('AC')->setWidth(14);
$excel->setActiveSheetIndex(0)->getColumnDimension('AD')->setWidth(12);
$excel->setActiveSheetIndex(0)->getColumnDimension('AE')->setWidth(13);

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

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('A'.$count,strtoupper($sender))
        -> getStyle('A'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('B'.$count,strtoupper($deltype))
        -> getStyle('B'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('C'.$count,'PQUE')
        -> getStyle('C'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C'.$count)
        -> applyFromArray($border);

//

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('D'.$count.':D'.$mergeNumber);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('D'.$count,strtoupper($month))
        -> getStyle('D'.$count.':D'.$mergeNumber)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('D'.$count.':D'.$mergeNumber)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('D'.$count.':D'.$mergeNumber)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('E'.$count.':E'.$mergeNumber);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('E'.$count,$cyclecode)
        -> getStyle('E'.$count.':E'.$mergeNumber)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('E'.$count.':E'.$mergeNumber)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('E'.$count.':E'.$mergeNumber)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('F'.$count.':F'.$mergeNumber);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('F'.$count,$pud)
        -> getStyle('F'.$count.':F'.$mergeNumber)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F'.$count.':F'.$mergeNumber)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F'.$count.':F'.$mergeNumber)
        -> applyFromArray($border);

//

$sql_total_pque = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_pque = mysqli_query($db_conn, $sql_total_pque);
$count_total_pque = mysqli_num_rows($query_total_pque);

$totalVolume = $totalVolume+$count_total_pque;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('G'.$count,$count_total_pque)
        -> getStyle('G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G'.$count)
        -> applyFromArray($border);

//

$timestamp = strtotime($pud);
$puddatetime = date("Y-m-d H:i:s", $timestamp);

$sql_total_pque_1day = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -1 DAY) AND record_status_status='1'";
$query_total_pque_1day = mysqli_query($db_conn, $sql_total_pque_1day);
$count_total_pque_1day = mysqli_num_rows($query_total_pque_1day);

$sql_total_pque_2day = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -2 DAY) AND record_status_status='1'";
$query_total_pque_2day = mysqli_query($db_conn, $sql_total_pque_2day);
$count_total_pque_2day = mysqli_num_rows($query_total_pque_2day);

$sql_total_pque_3day = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -3 DAY) AND record_status_status='1'";
$query_total_pque_3day = mysqli_query($db_conn, $sql_total_pque_3day);
$count_total_pque_3day = mysqli_num_rows($query_total_pque_3day);

$sql_total_pque_4day = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -4 DAY) AND record_status_status='1'";
$query_total_pque_4day = mysqli_query($db_conn, $sql_total_pque_4day);
$count_total_pque_4day = mysqli_num_rows($query_total_pque_4day);

$sql_total_pque_5day = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -5 DAY) AND record_status_status='1'";
$query_total_pque_5day = mysqli_query($db_conn, $sql_total_pque_5day);
$count_total_pque_5day = mysqli_num_rows($query_total_pque_5day);

$sql_total_pque_6day = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -6 DAY) AND record_status_status='1'";
$query_total_pque_6day = mysqli_query($db_conn, $sql_total_pque_6day);
$count_total_pque_6day = mysqli_num_rows($query_total_pque_6day);

$sql_total_pque_7day = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -7 DAY) AND record_status_status='1'";
$query_total_pque_7day = mysqli_query($db_conn, $sql_total_pque_7day);
$count_total_pque_7day = mysqli_num_rows($query_total_pque_7day);

$sql_total_pque_8day = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -8 DAY) AND record_status_status='1'";
$query_total_pque_8day = mysqli_query($db_conn, $sql_total_pque_8day);
$count_total_pque_8day = mysqli_num_rows($query_total_pque_8day);

$sql_total_pque_9day = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -9 DAY) AND record_status_status='1'";
$query_total_pque_9day = mysqli_query($db_conn, $sql_total_pque_9day);
$count_total_pque_9day = mysqli_num_rows($query_total_pque_9day);

$sql_total_pque_10day = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -10 DAY) AND record_status_status='1'";
$query_total_pque_10day = mysqli_query($db_conn, $sql_total_pque_10day);
$count_total_pque_10day = mysqli_num_rows($query_total_pque_10day);

$sql_total_pque_11day = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -11 DAY) AND record_status_status='1'";
$query_total_pque_11day = mysqli_query($db_conn, $sql_total_pque_11day);
$count_total_pque_11day = mysqli_num_rows($query_total_pque_11day);

$sql_total_pque_12day = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -12 DAY) AND record_status_status='1'";
$query_total_pque_12day = mysqli_query($db_conn, $sql_total_pque_12day);
$count_total_pque_12day = mysqli_num_rows($query_total_pque_12day);

$sql_total_pque_13day = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -13 DAY) AND record_status_status='1'";
$query_total_pque_13day = mysqli_query($db_conn, $sql_total_pque_13day);
$count_total_pque_13day = mysqli_num_rows($query_total_pque_13day);

$sql_total_pque_14day = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -14 DAY) AND record_status_status='1'";
$query_total_pque_14day = mysqli_query($db_conn, $sql_total_pque_14day);
$count_total_pque_14day = mysqli_num_rows($query_total_pque_14day);

$sql_total_pque_15day = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -15 DAY) AND record_status_status='1'";
$query_total_pque_15day = mysqli_query($db_conn, $sql_total_pque_15day);
$count_total_pque_15day = mysqli_num_rows($query_total_pque_15day);

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


$excel -> setActiveSheetIndex(0) 
        -> setCellValue('H'.$count,$count_total_pque_1day)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('I'.$count,$count_total_pque_2day)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('J'.$count,$count_total_pque_3day)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('K'.$count,$count_total_pque_4day)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('L'.$count,$count_total_pque_5day)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('M'.$count,$count_total_pque_6day)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('N'.$count,$count_total_pque_7day)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('O'.$count,$count_total_pque_8day)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('P'.$count,$count_total_pque_9day)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Q'.$count,$count_total_pque_10day)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('R'.$count,$count_total_pque_11day)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('S'.$count,$count_total_pque_12day)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('T'.$count,$count_total_pque_13day)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('U'.$count,$count_total_pque_14day)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('V'.$count,$count_total_pque_15day)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$sql_total_pque_delivered = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_status = 'DELIVERED' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_pque_delivered = mysqli_query($db_conn, $sql_total_pque_delivered);
$count_total_pque_delivered = mysqli_num_rows($query_total_pque_delivered);

$totalDelivered = $totalDelivered+$count_total_pque_delivered;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('W'.$count,$count_total_pque_delivered)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$sql_total_pque_rts = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_status = 'RTS' AND record_reason_rts NOT LIKE '%OUT OF SCOPE%' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_pque_rts = mysqli_query($db_conn, $sql_total_pque_rts);
$count_total_pque_rts = mysqli_num_rows($query_total_pque_rts);

$totalRTS = $totalRTS+$count_total_pque_rts;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('X'.$count,$count_total_pque_rts)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$sql_total_pque_oos = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_reason_rts LIKE '%OUT OF SCOPE%' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_pque_oos = mysqli_query($db_conn, $sql_total_pque_oos);
$count_total_pque_oos = mysqli_num_rows($query_total_pque_oos);

$totalOOS = $totalOOS+$count_total_pque_oos;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Y'.$count,$count_total_pque_oos)
        -> getStyle('Y'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y'.$count)
        -> applyFromArray($border);

$sql_total_pque_balance = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_status = '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_pque_balance = mysqli_query($db_conn, $sql_total_pque_balance);
$count_total_pque_balance = mysqli_num_rows($query_total_pque_balance);

$totalBalance = $totalBalance+$count_total_pque_balance;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Z'.$count,$count_total_pque_balance)
        -> getStyle('Z'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

if($count_total_pque_balance == 0) {
    $reportStatusPQUE = "Complete";
}

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AA'.$count,$reportStatusPQUE)
        -> getStyle('AA'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AA'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AA'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AB'.$count,'PQUE')
        -> getStyle('AB'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AB'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AB'.$count)
        -> applyFromArray($border);

$sql_total_pque1 = "SELECT * FROM ppbms_record WHERE record_area = 'PARANAQUE' AND record_status != '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_pque1 = mysqli_query($db_conn, $sql_total_pque1);
$count_total_pque1 = mysqli_num_rows($query_total_pque1);

$count_total_pque1 = $count_total_pque1-$count_total_pque_oos;

$totalVolumeInvoice = $totalVolumeInvoice+$count_total_pque1;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AC'.$count,$count_total_pque1)
        -> getStyle('AC'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AD'.$count,'9')
        -> getStyle('AD'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$amountPQUE = 9*$count_total_pque1-$count_total_pque_oos;

$totalAmount = $totalAmount+$amountPQUE;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AE'.$count,'Php'.number_format($amountPQUE))
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);

//2ND LINE

$count = $count+1;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('A'.$count,strtoupper($sender))
        -> getStyle('A'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('B'.$count,strtoupper($deltype))
        -> getStyle('B'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('C'.$count,'LPNAS')
        -> getStyle('C'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C'.$count)
        -> applyFromArray($border);

$sql_total_lpnas = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_lpnas = mysqli_query($db_conn, $sql_total_lpnas);
$count_total_lpnas = mysqli_num_rows($query_total_lpnas);

$totalVolume = $totalVolume+$count_total_lpnas;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('G'.$count,$count_total_lpnas)
        -> getStyle('G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G'.$count)
        -> applyFromArray($border);

$sql_total_lpnas_1day = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -1 DAY) AND record_status_status='1'";
$query_total_lpnas_1day = mysqli_query($db_conn, $sql_total_lpnas_1day);
$count_total_lpnas_1day = mysqli_num_rows($query_total_lpnas_1day);

$sql_total_lpnas_2day = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -2 DAY) AND record_status_status='1'";
$query_total_lpnas_2day = mysqli_query($db_conn, $sql_total_lpnas_2day);
$count_total_lpnas_2day = mysqli_num_rows($query_total_lpnas_2day);

$sql_total_lpnas_3day = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -3 DAY) AND record_status_status='1'";
$query_total_lpnas_3day = mysqli_query($db_conn, $sql_total_lpnas_3day);
$count_total_lpnas_3day = mysqli_num_rows($query_total_lpnas_3day);

$sql_total_lpnas_4day = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -4 DAY) AND record_status_status='1'";
$query_total_lpnas_4day = mysqli_query($db_conn, $sql_total_lpnas_4day);
$count_total_lpnas_4day = mysqli_num_rows($query_total_lpnas_4day);

$sql_total_lpnas_5day = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -5 DAY) AND record_status_status='1'";
$query_total_lpnas_5day = mysqli_query($db_conn, $sql_total_lpnas_5day);
$count_total_lpnas_5day = mysqli_num_rows($query_total_lpnas_5day);

$sql_total_lpnas_6day = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -6 DAY) AND record_status_status='1'";
$query_total_lpnas_6day = mysqli_query($db_conn, $sql_total_lpnas_6day);
$count_total_lpnas_6day = mysqli_num_rows($query_total_lpnas_6day);

$sql_total_lpnas_7day = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -7 DAY) AND record_status_status='1'";
$query_total_lpnas_7day = mysqli_query($db_conn, $sql_total_lpnas_7day);
$count_total_lpnas_7day = mysqli_num_rows($query_total_lpnas_7day);

$sql_total_lpnas_8day = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -8 DAY) AND record_status_status='1'";
$query_total_lpnas_8day = mysqli_query($db_conn, $sql_total_lpnas_8day);
$count_total_lpnas_8day = mysqli_num_rows($query_total_lpnas_8day);

$sql_total_lpnas_9day = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -9 DAY) AND record_status_status='1'";
$query_total_lpnas_9day = mysqli_query($db_conn, $sql_total_lpnas_9day);
$count_total_lpnas_9day = mysqli_num_rows($query_total_lpnas_9day);

$sql_total_lpnas_10day = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -10 DAY) AND record_status_status='1'";
$query_total_lpnas_10day = mysqli_query($db_conn, $sql_total_lpnas_10day);
$count_total_lpnas_10day = mysqli_num_rows($query_total_lpnas_10day);

$sql_total_lpnas_11day = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -11 DAY) AND record_status_status='1'";
$query_total_lpnas_11day = mysqli_query($db_conn, $sql_total_lpnas_11day);
$count_total_lpnas_11day = mysqli_num_rows($query_total_lpnas_11day);

$sql_total_lpnas_12day = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -12 DAY) AND record_status_status='1'";
$query_total_lpnas_12day = mysqli_query($db_conn, $sql_total_lpnas_12day);
$count_total_lpnas_12day = mysqli_num_rows($query_total_lpnas_12day);

$sql_total_lpnas_13day = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -13 DAY) AND record_status_status='1'";
$query_total_lpnas_13day = mysqli_query($db_conn, $sql_total_lpnas_13day);
$count_total_lpnas_13day = mysqli_num_rows($query_total_lpnas_13day);

$sql_total_lpnas_14day = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -14 DAY) AND record_status_status='1'";
$query_total_lpnas_14day = mysqli_query($db_conn, $sql_total_lpnas_14day);
$count_total_lpnas_14day = mysqli_num_rows($query_total_lpnas_14day);

$sql_total_lpnas_15day = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -15 DAY) AND record_status_status='1'";
$query_total_lpnas_15day = mysqli_query($db_conn, $sql_total_lpnas_15day);
$count_total_lpnas_15day = mysqli_num_rows($query_total_lpnas_15day);

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

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('H'.$count,$count_total_lpnas_1day)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('I'.$count,$count_total_lpnas_2day)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('J'.$count,$count_total_lpnas_3day)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('K'.$count,$count_total_lpnas_4day)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('L'.$count,$count_total_lpnas_5day)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('M'.$count,$count_total_lpnas_6day)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('N'.$count,$count_total_lpnas_7day)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('O'.$count,$count_total_lpnas_8day)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('P'.$count,$count_total_lpnas_9day)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Q'.$count,$count_total_lpnas_10day)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('R'.$count,$count_total_lpnas_11day)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('S'.$count,$count_total_lpnas_12day)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('T'.$count,$count_total_lpnas_13day)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('U'.$count,$count_total_lpnas_14day)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('V'.$count,$count_total_lpnas_15day)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$sql_total_lpnas_delivered = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_status = 'DELIVERED' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_lpnas_delivered = mysqli_query($db_conn, $sql_total_lpnas_delivered);
$count_total_lpnas_delivered = mysqli_num_rows($query_total_lpnas_delivered);

$totalDelivered = $totalDelivered+$count_total_lpnas_delivered;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('W'.$count,$count_total_lpnas_delivered)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$sql_total_lpnas_rts = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_status = 'RTS' AND record_reason_rts NOT LIKE '%OUT OF SCOPE%' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_lpnas_rts = mysqli_query($db_conn, $sql_total_lpnas_rts);
$count_total_lpnas_rts = mysqli_num_rows($query_total_lpnas_rts);

$totalRTS = $totalRTS+$count_total_lpnas_rts;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('X'.$count,$count_total_lpnas_rts)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$sql_total_lpnas_oos = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_reason_rts LIKE '%OUT OF SCOPE%' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_lpnas_oos = mysqli_query($db_conn, $sql_total_lpnas_oos);
$count_total_lpnas_oos = mysqli_num_rows($query_total_lpnas_oos);

$totalOOS = $totalOOS+$count_total_lpnas_oos;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Y'.$count,$count_total_lpnas_oos)
        -> getStyle('Y'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y'.$count)
        -> applyFromArray($border);

$sql_total_lpnas_balance = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_status = '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_lpnas_balance = mysqli_query($db_conn, $sql_total_lpnas_balance);
$count_total_lpnas_balance = mysqli_num_rows($query_total_lpnas_balance);

$totalBalance = $totalBalance+$count_total_lpnas_balance;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Z'.$count,$count_total_lpnas_balance)
        -> getStyle('Z'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

if($count_total_lpnas_balance == 0) {
    $reportStatusLPNAS = "Complete";
} 

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AA'.$count,$reportStatusLPNAS)
        -> getStyle('AA'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AA'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AA'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AB'.$count,'LPNAS')
        -> getStyle('AB'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AB'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AB'.$count)
        -> applyFromArray($border);

$sql_total_lpnas1 = "SELECT * FROM ppbms_record WHERE (record_area = 'LASPINAS' OR record_area = 'LAS PINAS') AND record_status != '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_lpnas1 = mysqli_query($db_conn, $sql_total_lpnas1);
$count_total_lpnas1 = mysqli_num_rows($query_total_lpnas1);

$count_total_lpnas1 = $count_total_lpnas1-$count_total_lpnas_oos;

$totalVolumeInvoice = $totalVolumeInvoice+$count_total_lpnas1;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AC'.$count,$count_total_lpnas1)
        -> getStyle('AC'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AD'.$count,'9')
        -> getStyle('AD'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$amountLPNAS = 9*$count_total_lpnas1-$count_total_lpnas_oos;

$totalAmount = $totalAmount+$amountLPNAS;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AE'.$count,'Php'.number_format($amountLPNAS))
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);

// 3RD LINE

$count = $count+1;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('A'.$count,strtoupper($sender))
        -> getStyle('A'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('B'.$count,strtoupper($deltype))
        -> getStyle('B'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('C'.$count,'MLUPA')
        -> getStyle('C'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C'.$count)
        -> applyFromArray($border);

$sql_total_mlupa = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mlupa = mysqli_query($db_conn, $sql_total_mlupa);
$count_total_mlupa = mysqli_num_rows($query_total_mlupa);

$totalVolume = $totalVolume+$count_total_mlupa;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('G'.$count,$count_total_mlupa)
        -> getStyle('G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G'.$count)
        -> applyFromArray($border);

$sql_total_mlupa_1day = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -1 DAY) AND record_status_status='1'";
$query_total_mlupa_1day = mysqli_query($db_conn, $sql_total_mlupa_1day);
$count_total_mlupa_1day = mysqli_num_rows($query_total_mlupa_1day);

$sql_total_mlupa_2day = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -2 DAY) AND record_status_status='1'";
$query_total_mlupa_2day = mysqli_query($db_conn, $sql_total_mlupa_2day);
$count_total_mlupa_2day = mysqli_num_rows($query_total_mlupa_2day);

$sql_total_mlupa_3day = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -3 DAY) AND record_status_status='1'";
$query_total_mlupa_3day = mysqli_query($db_conn, $sql_total_mlupa_3day);
$count_total_mlupa_3day = mysqli_num_rows($query_total_mlupa_3day);

$sql_total_mlupa_4day = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -4 DAY) AND record_status_status='1'";
$query_total_mlupa_4day = mysqli_query($db_conn, $sql_total_mlupa_4day);
$count_total_mlupa_4day = mysqli_num_rows($query_total_mlupa_4day);

$sql_total_mlupa_5day = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -5 DAY) AND record_status_status='1'";
$query_total_mlupa_5day = mysqli_query($db_conn, $sql_total_mlupa_5day);
$count_total_mlupa_5day = mysqli_num_rows($query_total_mlupa_5day);

$sql_total_mlupa_6day = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -6 DAY) AND record_status_status='1'";
$query_total_mlupa_6day = mysqli_query($db_conn, $sql_total_mlupa_6day);
$count_total_mlupa_6day = mysqli_num_rows($query_total_mlupa_6day);

$sql_total_mlupa_7day = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -7 DAY) AND record_status_status='1'";
$query_total_mlupa_7day = mysqli_query($db_conn, $sql_total_mlupa_7day);
$count_total_mlupa_7day = mysqli_num_rows($query_total_mlupa_7day);

$sql_total_mlupa_8day = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -8 DAY) AND record_status_status='1'";
$query_total_mlupa_8day = mysqli_query($db_conn, $sql_total_mlupa_8day);
$count_total_mlupa_8day = mysqli_num_rows($query_total_mlupa_8day);

$sql_total_mlupa_9day = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -9 DAY) AND record_status_status='1'";
$query_total_mlupa_9day = mysqli_query($db_conn, $sql_total_mlupa_9day);
$count_total_mlupa_9day = mysqli_num_rows($query_total_mlupa_9day);

$sql_total_mlupa_10day = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -10 DAY) AND record_status_status='1'";
$query_total_mlupa_10day = mysqli_query($db_conn, $sql_total_mlupa_10day);
$count_total_mlupa_10day = mysqli_num_rows($query_total_mlupa_10day);

$sql_total_mlupa_11day = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -11 DAY) AND record_status_status='1'";
$query_total_mlupa_11day = mysqli_query($db_conn, $sql_total_mlupa_11day);
$count_total_mlupa_11day = mysqli_num_rows($query_total_mlupa_11day);

$sql_total_mlupa_12day = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -12 DAY) AND record_status_status='1'";
$query_total_mlupa_12day = mysqli_query($db_conn, $sql_total_mlupa_12day);
$count_total_mlupa_12day = mysqli_num_rows($query_total_mlupa_12day);

$sql_total_mlupa_13day = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -13 DAY) AND record_status_status='1'";
$query_total_mlupa_13day = mysqli_query($db_conn, $sql_total_mlupa_13day);
$count_total_mlupa_13day = mysqli_num_rows($query_total_mlupa_13day);

$sql_total_mlupa_14day = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -14 DAY) AND record_status_status='1'";
$query_total_mlupa_14day = mysqli_query($db_conn, $sql_total_mlupa_14day);
$count_total_mlupa_14day = mysqli_num_rows($query_total_mlupa_14day);

$sql_total_mlupa_15day = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -15 DAY) AND record_status_status='1'";
$query_total_mlupa_15day = mysqli_query($db_conn, $sql_total_mlupa_15day);
$count_total_mlupa_15day = mysqli_num_rows($query_total_mlupa_15day);

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

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('H'.$count,$count_total_mlupa_1day)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('I'.$count,$count_total_mlupa_2day)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('J'.$count,$count_total_mlupa_3day)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('K'.$count,$count_total_mlupa_4day)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('L'.$count,$count_total_mlupa_5day)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('M'.$count,$count_total_mlupa_6day)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('N'.$count,$count_total_mlupa_7day)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('O'.$count,$count_total_mlupa_8day)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('P'.$count,$count_total_mlupa_9day)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Q'.$count,$count_total_mlupa_10day)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('R'.$count,$count_total_mlupa_11day)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('S'.$count,$count_total_mlupa_12day)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('T'.$count,$count_total_mlupa_13day)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('U'.$count,$count_total_mlupa_14day)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('V'.$count,$count_total_mlupa_15day)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$sql_total_mlupa_delivered = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_status = 'DELIVERED' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mlupa_delivered = mysqli_query($db_conn, $sql_total_mlupa_delivered);
$count_total_mlupa_delivered = mysqli_num_rows($query_total_mlupa_delivered);

$totalDelivered = $totalDelivered+$count_total_mlupa_delivered;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('W'.$count,$count_total_mlupa_delivered)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$sql_total_mlupa_rts = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_status = 'RTS' AND record_reason_rts NOT LIKE '%OUT OF SCOPE%' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mlupa_rts = mysqli_query($db_conn, $sql_total_mlupa_rts);
$count_total_mlupa_rts = mysqli_num_rows($query_total_mlupa_rts);

$totalRTS = $totalRTS+$count_total_mlupa_rts;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('X'.$count,$count_total_mlupa_rts)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$sql_total_mlupa_oos = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_reason_rts LIKE '%OUT OF SCOPE%' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mlupa_oos = mysqli_query($db_conn, $sql_total_mlupa_oos);
$count_total_mlupa_oos = mysqli_num_rows($query_total_mlupa_oos);

$totalOOS = $totalOOS+$count_total_mlupa_oos;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Y'.$count,$count_total_mlupa_oos)
        -> getStyle('Y'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y'.$count)
        -> applyFromArray($border);

$sql_total_mlupa_balance = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_status = '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mlupa_balance = mysqli_query($db_conn, $sql_total_mlupa_balance);
$count_total_mlupa_balance = mysqli_num_rows($query_total_mlupa_balance);

$totalBalance = $totalBalance+$count_total_mlupa_balance;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Z'.$count,$count_total_mlupa_balance)
        -> getStyle('Z'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

if($count_total_mlupa_balance == 0) {
    $reportStatusMLUPA = "Complete";
}

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AA'.$count,$reportStatusMLUPA)
        -> getStyle('AA'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AA'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AA'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AB'.$count,'MLUPA')
        -> getStyle('AB'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AB'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AB'.$count)
        -> applyFromArray($border);

$sql_total_mlupa1 = "SELECT * FROM ppbms_record WHERE record_area = 'MUNTINLUPA' AND record_status != '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mlupa1 = mysqli_query($db_conn, $sql_total_mlupa1);
$count_total_mlupa1 = mysqli_num_rows($query_total_mlupa1);

$count_total_mlupa1 = $count_total_mlupa1-$count_total_mlupa_oos;

$totalVolumeInvoice = $totalVolumeInvoice+$count_total_mlupa1;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AC'.$count,$count_total_mlupa1)
        -> getStyle('AC'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AD'.$count,'9')
        -> getStyle('AD'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$amountMLUPA = 9*$count_total_mlupa1-$count_total_mlupa_oos;

$totalAmount = $totalAmount+$amountMLUPA;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AE'.$count,'Php'.number_format($amountMLUPA))
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);

// 4TH LINE

$count = $count+1;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('A'.$count,strtoupper($sender))
        -> getStyle('A'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('B'.$count,strtoupper($deltype))
        -> getStyle('B'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('C'.$count,'MKNA')
        -> getStyle('C'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C'.$count)
        -> applyFromArray($border);

$sql_total_mkna = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mkna = mysqli_query($db_conn, $sql_total_mkna);
$count_total_mkna = mysqli_num_rows($query_total_mkna);

$totalVolume = $totalVolume+$count_total_mkna;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('G'.$count,$count_total_mkna)
        -> getStyle('G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G'.$count)
        -> applyFromArray($border);

$sql_total_mkna_1day = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -1 DAY) AND record_status_status='1'";
$query_total_mkna_1day = mysqli_query($db_conn, $sql_total_mkna_1day);
$count_total_mkna_1day = mysqli_num_rows($query_total_mkna_1day);

$sql_total_mkna_2day = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -2 DAY) AND record_status_status='1'";
$query_total_mkna_2day = mysqli_query($db_conn, $sql_total_mkna_2day);
$count_total_mkna_2day = mysqli_num_rows($query_total_mkna_2day);

$sql_total_mkna_3day = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -3 DAY) AND record_status_status='1'";
$query_total_mkna_3day = mysqli_query($db_conn, $sql_total_mkna_3day);
$count_total_mkna_3day = mysqli_num_rows($query_total_mkna_3day);

$sql_total_mkna_4day = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -4 DAY) AND record_status_status='1'";
$query_total_mkna_4day = mysqli_query($db_conn, $sql_total_mkna_4day);
$count_total_mkna_4day = mysqli_num_rows($query_total_mkna_4day);

$sql_total_mkna_5day = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -5 DAY) AND record_status_status='1'";
$query_total_mkna_5day = mysqli_query($db_conn, $sql_total_mkna_5day);
$count_total_mkna_5day = mysqli_num_rows($query_total_mkna_5day);

$sql_total_mkna_6day = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -6 DAY) AND record_status_status='1'";
$query_total_mkna_6day = mysqli_query($db_conn, $sql_total_mkna_6day);
$count_total_mkna_6day = mysqli_num_rows($query_total_mkna_6day);

$sql_total_mkna_7day = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -7 DAY) AND record_status_status='1'";
$query_total_mkna_7day = mysqli_query($db_conn, $sql_total_mkna_7day);
$count_total_mkna_7day = mysqli_num_rows($query_total_mkna_7day);

$sql_total_mkna_8day = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -8 DAY) AND record_status_status='1'";
$query_total_mkna_8day = mysqli_query($db_conn, $sql_total_mkna_8day);
$count_total_mkna_8day = mysqli_num_rows($query_total_mkna_8day);

$sql_total_mkna_9day = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -9 DAY) AND record_status_status='1'";
$query_total_mkna_9day = mysqli_query($db_conn, $sql_total_mkna_9day);
$count_total_mkna_9day = mysqli_num_rows($query_total_mkna_9day);

$sql_total_mkna_10day = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -10 DAY) AND record_status_status='1'";
$query_total_mkna_10day = mysqli_query($db_conn, $sql_total_mkna_10day);
$count_total_mkna_10day = mysqli_num_rows($query_total_mkna_10day);

$sql_total_mkna_11day = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -11 DAY) AND record_status_status='1'";
$query_total_mkna_11day = mysqli_query($db_conn, $sql_total_mkna_11day);
$count_total_mkna_11day = mysqli_num_rows($query_total_mkna_11day);

$sql_total_mkna_12day = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -12 DAY) AND record_status_status='1'";
$query_total_mkna_12day = mysqli_query($db_conn, $sql_total_mkna_12day);
$count_total_mkna_12day = mysqli_num_rows($query_total_mkna_12day);

$sql_total_mkna_13day = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -13 DAY) AND record_status_status='1'";
$query_total_mkna_13day = mysqli_query($db_conn, $sql_total_mkna_13day);
$count_total_mkna_13day = mysqli_num_rows($query_total_mkna_13day);

$sql_total_mkna_14day = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -14 DAY) AND record_status_status='1'";
$query_total_mkna_14day = mysqli_query($db_conn, $sql_total_mkna_14day);
$count_total_mkna_14day = mysqli_num_rows($query_total_mkna_14day);

$sql_total_mkna_15day = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -15 DAY) AND record_status_status='1'";
$query_total_mkna_15day = mysqli_query($db_conn, $sql_total_mkna_15day);
$count_total_mkna_15day = mysqli_num_rows($query_total_mkna_15day);

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

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('H'.$count,$count_total_mkna_1day)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('I'.$count,$count_total_mkna_2day)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('J'.$count,$count_total_mkna_3day)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('K'.$count,$count_total_mkna_4day)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('L'.$count,$count_total_mkna_5day)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('M'.$count,$count_total_mkna_6day)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('N'.$count,$count_total_mkna_7day)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('O'.$count,$count_total_mkna_8day)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('P'.$count,$count_total_mkna_9day)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Q'.$count,$count_total_mkna_10day)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('R'.$count,$count_total_mkna_11day)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('S'.$count,$count_total_mkna_12day)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('T'.$count,$count_total_mkna_13day)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('U'.$count,$count_total_mkna_14day)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('V'.$count,$count_total_mkna_15day)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$sql_total_mkna_delivered = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_status = 'DELIVERED' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mkna_delivered = mysqli_query($db_conn, $sql_total_mkna_delivered);
$count_total_mkna_delivered = mysqli_num_rows($query_total_mkna_delivered);

$totalDelivered = $totalDelivered+$count_total_mkna_delivered;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('W'.$count,$count_total_mkna_delivered)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$sql_total_mkna_rts = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_status = 'RTS' AND record_reason_rts NOT LIKE '%OUT OF SCOPE%' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mkna_rts = mysqli_query($db_conn, $sql_total_mkna_rts);
$count_total_mkna_rts = mysqli_num_rows($query_total_mkna_rts);

$totalRTS = $totalRTS+$count_total_mkna_rts;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('X'.$count,$count_total_mkna_rts)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$sql_total_mkna_oos = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_reason_rts LIKE '%OUT OF SCOPE%' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mkna_oos = mysqli_query($db_conn, $sql_total_mkna_oos);
$count_total_mkna_oos = mysqli_num_rows($query_total_mkna_oos);

$totalOOS = $totalOOS+$count_total_mkna_oos;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Y'.$count,$count_total_mkna_oos)
        -> getStyle('Y'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y'.$count)
        -> applyFromArray($border);

$sql_total_mkna_balance = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_status = '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mkna_balance = mysqli_query($db_conn, $sql_total_mkna_balance);
$count_total_mkna_balance = mysqli_num_rows($query_total_mkna_balance);

$totalBalance = $totalBalance+$count_total_mkna_balance;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Z'.$count,$count_total_mkna_balance)
        -> getStyle('Z'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

if($count_total_mkna_balance == 0) {
    $reportStatusMKNA = "Complete";
}

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AA'.$count,$reportStatusMKNA)
        -> getStyle('AA'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AA'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AA'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AB'.$count,'MKNA')
        -> getStyle('AB'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AB'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AB'.$count)
        -> applyFromArray($border);

$sql_total_mkna1 = "SELECT * FROM ppbms_record WHERE record_area = 'MARIKINA' AND record_status != '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_mkna1 = mysqli_query($db_conn, $sql_total_mkna1);
$count_total_mkna1 = mysqli_num_rows($query_total_mkna1);

$count_total_mkna1 = $count_total_mkna1-$count_total_mkna_oos;

$totalVolumeInvoice = $totalVolumeInvoice+$count_total_mkna1;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AC'.$count,$count_total_mkna1)
        -> getStyle('AC'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AD'.$count,'9')
        -> getStyle('AD'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$amountMKNA = 9*$count_total_mkna1-$count_total_mkna_oos;

$totalAmount = $totalAmount+$amountMKNA;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AE'.$count,'Php'.number_format($amountMKNA))
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);

//5TH LINE

$count = $count+1;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('A'.$count,strtoupper($sender))
        -> getStyle('A'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('B'.$count,strtoupper($deltype))
        -> getStyle('B'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('C'.$count,'RIZAL')
        -> getStyle('C'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C'.$count)
        -> applyFromArray($border);

$sql_total_rizal = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_rizal = mysqli_query($db_conn, $sql_total_rizal);
$count_total_rizal = mysqli_num_rows($query_total_rizal);

$totalVolume = $totalVolume+$count_total_rizal;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('G'.$count,$count_total_rizal)
        -> getStyle('G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G'.$count)
        -> applyFromArray($border);

$sql_total_rizal_1day = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -1 DAY) AND record_status_status='1'";
$query_total_rizal_1day = mysqli_query($db_conn, $sql_total_rizal_1day);
$count_total_rizal_1day = mysqli_num_rows($query_total_rizal_1day);

$sql_total_rizal_2day = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -2 DAY) AND record_status_status='1'";
$query_total_rizal_2day = mysqli_query($db_conn, $sql_total_rizal_2day);
$count_total_rizal_2day = mysqli_num_rows($query_total_rizal_2day);

$sql_total_rizal_3day = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -3 DAY) AND record_status_status='1'";
$query_total_rizal_3day = mysqli_query($db_conn, $sql_total_rizal_3day);
$count_total_rizal_3day = mysqli_num_rows($query_total_rizal_3day);

$sql_total_rizal_4day = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -4 DAY) AND record_status_status='1'";
$query_total_rizal_4day = mysqli_query($db_conn, $sql_total_rizal_4day);
$count_total_rizal_4day = mysqli_num_rows($query_total_rizal_4day);

$sql_total_rizal_5day = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -5 DAY) AND record_status_status='1'";
$query_total_rizal_5day = mysqli_query($db_conn, $sql_total_rizal_5day);
$count_total_rizal_5day = mysqli_num_rows($query_total_rizal_5day);

$sql_total_rizal_6day = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -6 DAY) AND record_status_status='1'";
$query_total_rizal_6day = mysqli_query($db_conn, $sql_total_rizal_6day);
$count_total_rizal_6day = mysqli_num_rows($query_total_rizal_6day);

$sql_total_rizal_7day = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -7 DAY) AND record_status_status='1'";
$query_total_rizal_7day = mysqli_query($db_conn, $sql_total_rizal_7day);
$count_total_rizal_7day = mysqli_num_rows($query_total_rizal_7day);

$sql_total_rizal_8day = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -8 DAY) AND record_status_status='1'";
$query_total_rizal_8day = mysqli_query($db_conn, $sql_total_rizal_8day);
$count_total_rizal_8day = mysqli_num_rows($query_total_rizal_8day);

$sql_total_rizal_9day = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -9 DAY) AND record_status_status='1'";
$query_total_rizal_9day = mysqli_query($db_conn, $sql_total_rizal_9day);
$count_total_rizal_9day = mysqli_num_rows($query_total_rizal_9day);

$sql_total_rizal_10day = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -10 DAY) AND record_status_status='1'";
$query_total_rizal_10day = mysqli_query($db_conn, $sql_total_rizal_10day);
$count_total_rizal_10day = mysqli_num_rows($query_total_rizal_10day);

$sql_total_rizal_11day = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -11 DAY) AND record_status_status='1'";
$query_total_rizal_11day = mysqli_query($db_conn, $sql_total_rizal_11day);
$count_total_rizal_11day = mysqli_num_rows($query_total_rizal_11day);

$sql_total_rizal_12day = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -12 DAY) AND record_status_status='1'";
$query_total_rizal_12day = mysqli_query($db_conn, $sql_total_rizal_12day);
$count_total_rizal_12day = mysqli_num_rows($query_total_rizal_12day);

$sql_total_rizal_13day = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -13 DAY) AND record_status_status='1'";
$query_total_rizal_13day = mysqli_query($db_conn, $sql_total_rizal_13day);
$count_total_rizal_13day = mysqli_num_rows($query_total_rizal_13day);

$sql_total_rizal_14day = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -14 DAY) AND record_status_status='1'";
$query_total_rizal_14day = mysqli_query($db_conn, $sql_total_rizal_14day);
$count_total_rizal_14day = mysqli_num_rows($query_total_rizal_14day);

$sql_total_rizal_15day = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -15 DAY) AND record_status_status='1'";
$query_total_rizal_15day = mysqli_query($db_conn, $sql_total_rizal_15day);
$count_total_rizal_15day = mysqli_num_rows($query_total_rizal_15day);

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

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('H'.$count,$count_total_rizal_1day)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('I'.$count,$count_total_rizal_2day)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('J'.$count,$count_total_rizal_3day)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('K'.$count,$count_total_rizal_4day)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('L'.$count,$count_total_rizal_5day)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('M'.$count,$count_total_rizal_6day)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('N'.$count,$count_total_rizal_7day)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('O'.$count,$count_total_rizal_8day)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('P'.$count,$count_total_rizal_9day)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Q'.$count,$count_total_rizal_10day)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('R'.$count,$count_total_rizal_11day)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('S'.$count,$count_total_rizal_12day)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('T'.$count,$count_total_rizal_13day)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('U'.$count,$count_total_rizal_14day)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('V'.$count,$count_total_rizal_15day)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$sql_total_rizal_delivered = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_status = 'DELIVERED' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_rizal_delivered = mysqli_query($db_conn, $sql_total_rizal_delivered);
$count_total_rizal_delivered = mysqli_num_rows($query_total_rizal_delivered);

$totalDelivered = $totalDelivered+$count_total_rizal_delivered;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('W'.$count,$count_total_rizal_delivered)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$sql_total_rizal_rts = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_status = 'RTS' AND record_reason_rts NOT LIKE '%OUT OF SCOPE%' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_rizal_rts = mysqli_query($db_conn, $sql_total_rizal_rts);
$count_total_rizal_rts = mysqli_num_rows($query_total_rizal_rts);

$totalRTS = $totalRTS+$count_total_rizal_rts;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('X'.$count,$count_total_rizal_rts)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$sql_total_rizal_oos = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_reason_rts LIKE '%OUT OF SCOPE%' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_rizal_oos = mysqli_query($db_conn, $sql_total_rizal_oos);
$count_total_rizal_oos = mysqli_num_rows($query_total_rizal_oos);

$totalOOS = $totalOOS+$count_total_rizal_oos;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Y'.$count,$count_total_rizal_oos)
        -> getStyle('Y'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y'.$count)
        -> applyFromArray($border);

$sql_total_rizal_balance = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_status = '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_rizal_balance = mysqli_query($db_conn, $sql_total_rizal_balance);
$count_total_rizal_balance = mysqli_num_rows($query_total_rizal_balance);

$totalBalance = $totalBalance+$count_total_rizal_balance;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Z'.$count,$count_total_rizal_balance)
        -> getStyle('Z'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

if($count_total_rizal_balance == 0) {
    $reportStatusRIZAL = "Complete";
}

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AA'.$count,$reportStatusRIZAL)
        -> getStyle('AA'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AA'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AA'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AB'.$count,'RIZAL')
        -> getStyle('AB'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AB'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AB'.$count)
        -> applyFromArray($border);

$sql_total_rizal1 = "SELECT * FROM ppbms_record WHERE record_area = 'RIZAL' AND record_status != '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_rizal1 = mysqli_query($db_conn, $sql_total_rizal1);
$count_total_rizal1 = mysqli_num_rows($query_total_rizal1);

$count_total_rizal1 = $count_total_rizal1-$count_total_rizal_oos;

$totalVolumeInvoice = $totalVolumeInvoice+$count_total_rizal1;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AC'.$count,$count_total_rizal1)
        -> getStyle('AC'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AD'.$count,'10')
        -> getStyle('AD'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$amountRIZAL = 9*$count_total_rizal1-$count_total_rizal_oos;

$totalAmount = $totalAmount+$amountRIZAL;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AE'.$count,'Php'.number_format($amountRIZAL))
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);

//6TH LINE

$count = $count+1;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('A'.$count,strtoupper($sender))
        -> getStyle('A'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('A'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('B'.$count,strtoupper($deltype))
        -> getStyle('B'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('B'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('C'.$count,'LAGUNA')
        -> getStyle('C'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('C'.$count)
        -> applyFromArray($border);

$sql_total_laguna = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_laguna = mysqli_query($db_conn, $sql_total_laguna);
$count_total_laguna = mysqli_num_rows($query_total_laguna);

$totalVolume = $totalVolume+$count_total_laguna;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('G'.$count,$count_total_laguna)
        -> getStyle('G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G'.$count)
        -> applyFromArray($border);

$sql_total_laguna_1day = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -1 DAY) AND record_status_status='1'";
$query_total_laguna_1day = mysqli_query($db_conn, $sql_total_laguna_1day);
$count_total_laguna_1day = mysqli_num_rows($query_total_laguna_1day);

$sql_total_laguna_2day = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -2 DAY) AND record_status_status='1'";
$query_total_laguna_2day = mysqli_query($db_conn, $sql_total_laguna_2day);
$count_total_laguna_2day = mysqli_num_rows($query_total_laguna_2day);

$sql_total_laguna_3day = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -3 DAY) AND record_status_status='1'";
$query_total_laguna_3day = mysqli_query($db_conn, $sql_total_laguna_3day);
$count_total_laguna_3day = mysqli_num_rows($query_total_laguna_3day);

$sql_total_laguna_4day = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -4 DAY) AND record_status_status='1'";
$query_total_laguna_4day = mysqli_query($db_conn, $sql_total_laguna_4day);
$count_total_laguna_4day = mysqli_num_rows($query_total_laguna_4day);

$sql_total_laguna_5day = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -5 DAY) AND record_status_status='1'";
$query_total_laguna_5day = mysqli_query($db_conn, $sql_total_laguna_5day);
$count_total_laguna_5day = mysqli_num_rows($query_total_laguna_5day);

$sql_total_laguna_6day = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -6 DAY) AND record_status_status='1'";
$query_total_laguna_6day = mysqli_query($db_conn, $sql_total_laguna_6day);
$count_total_laguna_6day = mysqli_num_rows($query_total_laguna_6day);

$sql_total_laguna_7day = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -7 DAY) AND record_status_status='1'";
$query_total_laguna_7day = mysqli_query($db_conn, $sql_total_laguna_7day);
$count_total_laguna_7day = mysqli_num_rows($query_total_laguna_7day);

$sql_total_laguna_8day = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -8 DAY) AND record_status_status='1'";
$query_total_laguna_8day = mysqli_query($db_conn, $sql_total_laguna_8day);
$count_total_laguna_8day = mysqli_num_rows($query_total_laguna_8day);

$sql_total_laguna_9day = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -9 DAY) AND record_status_status='1'";
$query_total_laguna_9day = mysqli_query($db_conn, $sql_total_laguna_9day);
$count_total_laguna_9day = mysqli_num_rows($query_total_laguna_9day);

$sql_total_laguna_10day = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -10 DAY) AND record_status_status='1'";
$query_total_laguna_10day = mysqli_query($db_conn, $sql_total_laguna_10day);
$count_total_laguna_10day = mysqli_num_rows($query_total_laguna_10day);

$sql_total_laguna_11day = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -11 DAY) AND record_status_status='1'";
$query_total_laguna_11day = mysqli_query($db_conn, $sql_total_laguna_11day);
$count_total_laguna_11day = mysqli_num_rows($query_total_laguna_11day);

$sql_total_laguna_12day = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -12 DAY) AND record_status_status='1'";
$query_total_laguna_12day = mysqli_query($db_conn, $sql_total_laguna_12day);
$count_total_laguna_12day = mysqli_num_rows($query_total_laguna_12day);

$sql_total_laguna_13day = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -13 DAY) AND record_status_status='1'";
$query_total_laguna_13day = mysqli_query($db_conn, $sql_total_laguna_13day);
$count_total_laguna_13day = mysqli_num_rows($query_total_laguna_13day);

$sql_total_laguna_14day = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -14 DAY) AND record_status_status='1'";
$query_total_laguna_14day = mysqli_query($db_conn, $sql_total_laguna_14day);
$count_total_laguna_14day = mysqli_num_rows($query_total_laguna_14day);

$sql_total_laguna_15day = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_date_receive = DATE_SUB('$puddatetime', INTERVAL -15 DAY) AND record_status_status='1'";
$query_total_laguna_15day = mysqli_query($db_conn, $sql_total_laguna_15day);
$count_total_laguna_15day = mysqli_num_rows($query_total_laguna_15day);

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

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('H'.$count,$count_total_laguna_1day)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('I'.$count,$count_total_laguna_2day)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('J'.$count,$count_total_laguna_3day)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('K'.$count,$count_total_laguna_4day)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('L'.$count,$count_total_laguna_5day)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('M'.$count,$count_total_laguna_6day)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('N'.$count,$count_total_laguna_7day)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('O'.$count,$count_total_laguna_8day)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('P'.$count,$count_total_laguna_9day)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Q'.$count,$count_total_laguna_10day)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('R'.$count,$count_total_laguna_11day)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('S'.$count,$count_total_laguna_12day)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('T'.$count,$count_total_laguna_13day)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('U'.$count,$count_total_laguna_14day)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('V'.$count,$count_total_laguna_15day)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$sql_total_laguna_delivered = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_status = 'DELIVERED' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_laguna_delivered = mysqli_query($db_conn, $sql_total_laguna_delivered);
$count_total_laguna_delivered = mysqli_num_rows($query_total_laguna_delivered);

$totalDelivered = $totalDelivered+$count_total_laguna_delivered;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('W'.$count,$count_total_laguna_delivered)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$sql_total_laguna_rts = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_status = 'RTS' AND record_reason_rts NOT LIKE '%OUT OF SCOPE%' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_laguna_rts = mysqli_query($db_conn, $sql_total_laguna_rts);
$count_total_laguna_rts = mysqli_num_rows($query_total_laguna_rts);

$totalRTS = $totalRTS+$count_total_laguna_rts;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('X'.$count,$count_total_laguna_rts)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$sql_total_laguna_oos = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_reason_rts LIKE '%OUT OF SCOPE%' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_laguna_oos = mysqli_query($db_conn, $sql_total_laguna_oos);
$count_total_laguna_oos = mysqli_num_rows($query_total_laguna_oos);

$totalOOS = $totalOOS+$count_total_laguna_oos;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Y'.$count,$count_total_laguna_oos)
        -> getStyle('Y'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y'.$count)
        -> applyFromArray($border);

$sql_total_laguna_balance = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_status = '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_laguna_balance = mysqli_query($db_conn, $sql_total_laguna_balance);
$count_total_laguna_balance = mysqli_num_rows($query_total_laguna_balance);

$totalBalance = $totalBalance+$count_total_laguna_balance;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Z'.$count,$count_total_laguna_balance)
        -> getStyle('Z'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

if($count_total_laguna_balance == 0) {
    $reportStatusLAGUNA = "Complete";
}

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AA'.$count,$reportStatusLAGUNA)
        -> getStyle('AA'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AA'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AA'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AB'.$count,'LAGUNA')
        -> getStyle('AB'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AB'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AB'.$count)
        -> applyFromArray($border);

$sql_total_laguna1 = "SELECT * FROM ppbms_record WHERE record_area = 'LAGUNA' AND record_status != '' AND record_cycle_code='$cyclecode' AND record_pud='$pud' AND record_status_status='1'";
$query_total_laguna1 = mysqli_query($db_conn, $sql_total_laguna1);
$count_total_laguna1 = mysqli_num_rows($query_total_laguna1);

$count_total_laguna1 = $count_total_laguna1-$count_total_laguna_oos;

$totalVolumeInvoice = $totalVolumeInvoice+$count_total_laguna1;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AC'.$count,$count_total_laguna1)
        -> getStyle('AC'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AD'.$count,'10')
        -> getStyle('AD'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$amountLAGUNA = 9*$count_total_laguna1-$count_total_laguna_oos;

$totalAmount = $totalAmount+$amountLAGUNA;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AE'.$count,'Php'.number_format($amountLAGUNA))
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);


$count = $count+1;

}

// FOOTER

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('F'.$count,'TOTAL')
        -> getStyle('F'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('G'.$count,$totalVolume)
        -> getStyle('G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('G'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('H'.$count,$totalDay1)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('I'.$count,$totalDay2)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('J'.$count,$totalDay3)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('K'.$count,$totalDay4)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('L'.$count,$totalDay5)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('M'.$count,$totalDay6)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('N'.$count,$totalDay7)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('O'.$count,$totalDay8)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('P'.$count,$totalDay9)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Q'.$count,$totalDay10)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('R'.$count,$totalDay11)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('S'.$count,$totalDay12)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('T'.$count,$totalDay13)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('U'.$count,$totalDay14)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('V'.$count,$totalDay15)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('W'.$count,$totalDelivered)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('X'.$count,$totalRTS)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Y'.$count,$totalOOS)
        -> getStyle('Y'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Z'.$count,$totalBalance)
        -> getStyle('Z'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AA'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('AB'.$count.':AE'.$count);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AB'.$count.':AE'.$count)
        -> applyFromArray($border);

$count = $count+1;

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('F'.$count.':G'.$count);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('F'.$count,'PERCENTAGE')
        -> getStyle('F'.$count.':G'.$count)
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
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

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('H'.$count,$day1Percent.'%')
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('I'.$count,$day2Percent.'%')
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('J'.$count,$day3Percent.'%')
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('K'.$count,$day4Percent.'%')
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('L'.$count,$day5Percent.'%')
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('M'.$count,$day6Percent.'%')
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('N'.$count,$day7Percent.'%')
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('O'.$count,$day8Percent.'%')
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('P'.$count,$day9Percent.'%')
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Q'.$count,$day10Percent.'%')
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('R'.$count,$day11Percent.'%')
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('S'.$count,$day12Percent.'%')
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('T'.$count,$day13Percent.'%')
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('U'.$count,$day14Percent.'%')
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('V'.$count,$day15Percent.'%')
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$deliveredPercent = get_percentage($totalVolume, $totalDelivered);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('W'.$count,$deliveredPercent.'%')
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$rtsPercent = get_percentage($totalVolume, $totalRTS);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('X'.$count,$rtsPercent.'%')
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$oosPercent = get_percentage($totalVolume, $totalOOS);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Y'.$count,$oosPercent.'%')
        -> getStyle('Y'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Y'.$count)
        -> applyFromArray($border);

$balancePercent = get_percentage($totalVolume, $totalBalance);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Z'.$count,$balancePercent.'%')
        -> getStyle('Z'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Z'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AA'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AB'.$count)
        -> applyFromArray($border);

$volumInvoicePercent = get_percentage($totalVolume, $totalVolumeInvoice);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AC'.$count,$volumInvoicePercent)
        -> getStyle('AC'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AC'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AE'.$count,'Php'.number_format($totalAmount, 2))
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);

$count = $count+1;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AD'.$count,'ADD 12% Vat')
        -> getStyle('AD'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$vatAmount = $totalAmount*0.12;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AE'.$count,$vatAmount)
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);

$count = $count+1;

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('F'.$count.':V'.$count);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F'.$count.':V'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F'.$count.':V'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('366092');

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('W'.$count.':X'.$count);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('W'.$count,'Total')
        -> getStyle('W'.$count.':X'.$count)
        -> applyFromArray($tableHeaderStyleFontYellowCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count.':X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count.':X'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count.':X'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('366092');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AD'.$count,'TOTAL')
        -> getStyle('AD'.$count)
        -> applyFromArray($tableHeaderStyleFontWhiteCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AD'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$grandTotalAmount = $totalAmount+$vatAmount;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('AE'.$count,$grandTotalAmount)
        -> getStyle('AE'.$count)
        -> applyFromArray($tableHeaderStyleFontWhiteBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('AE'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('000000');

$count = $count+1;

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('F'.$count.':V'.$count);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F'.$count.':V'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F'.$count.':V'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('W'.$count,'On Time')
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('X'.$count,'Late')
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontRedCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('ffff00');

$count = $count+1;

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('F'.$count.':G'.$count);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('F'.$count,'Provincial Delivery Status')
        -> getStyle('F'.$count.':G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F'.$count.':G'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('H'.$count,$provTotalDay1)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('I'.$count,$provTotalDay2)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('J'.$count,$provTotalDay3)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('K'.$count,$provTotalDay4)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('L'.$count,$provTotalDay5)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('M'.$count,$provTotalDay6)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('N'.$count,$provTotalDay7)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('O'.$count,$provTotalDay8)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('P'.$count,$provTotalDay9)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Q'.$count,$provTotalDay10)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('R'.$count,$provTotalDay11)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('S'.$count,$provTotalDay12)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('T'.$count,$provTotalDay13)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('U'.$count,$provTotalDay14)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('V'.$count,$provTotalDay15)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$totalNCROnTime = $ncrTotalDay1+$ncrTotalDay2+$ncrTotalDay3+$ncrTotalDay4+$ncrTotalDay5+$ncrTotalDay6+$ncrTotalDay7+$ncrTotalDay8+$ncrTotalDay9+$ncrTotalDay10;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('W'.$count,$totalNCROnTime)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$totalNCRLate = $ncrTotalDay11+$ncrTotalDay12+$ncrTotalDay13+$ncrTotalDay14+$ncrTotalDay15;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('X'.$count,$totalNCRLate)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$count = $count+1;

$excel -> setActiveSheetIndex(0) 
        -> mergeCells('F'.$count.':G'.$count);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('F'.$count,'NCR Delivery Status')
        -> getStyle('F'.$count.':G'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('F'.$count.':G'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('H'.$count,$ncrTotalDay1)
        -> getStyle('H'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('H'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('I'.$count,$ncrTotalDay2)
        -> getStyle('I'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('I'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('J'.$count,$ncrTotalDay3)
        -> getStyle('J'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('J'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('K'.$count,$ncrTotalDay4)
        -> getStyle('K'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('K'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('L'.$count,$ncrTotalDay5)
        -> getStyle('L'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('L'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('M'.$count,$ncrTotalDay6)
        -> getStyle('M'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('M'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('N'.$count,$ncrTotalDay7)
        -> getStyle('N'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('N'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('O'.$count,$ncrTotalDay8)
        -> getStyle('O'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('O'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('P'.$count,$ncrTotalDay9)
        -> getStyle('P'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('P'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('Q'.$count,$ncrTotalDay10)
        -> getStyle('Q'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('Q'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('R'.$count,$ncrTotalDay11)
        -> getStyle('R'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('R'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('S'.$count,$ncrTotalDay12)
        -> getStyle('S'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('S'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('T'.$count,$ncrTotalDay13)
        -> getStyle('T'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('T'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('U'.$count,$ncrTotalDay14)
        -> getStyle('U'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('U'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('V'.$count,$ncrTotalDay15)
        -> getStyle('V'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('V'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$totalProvOnTime = $provTotalDay1+$provTotalDay2+$provTotalDay3+$provTotalDay4+$provTotalDay5;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('W'.$count,$totalProvOnTime)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$totalProvLate = $provTotalDay6+$provTotalDay7+$provTotalDay8+$provTotalDay9+$provTotalDay10+$provTotalDay11+$provTotalDay12+$provTotalDay13+$provTotalDay14+$provTotalDay15;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('X'.$count,$totalProvLate)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('c5d9f1');

$totalOnTime = $totalNCROnTime+$totalProvOnTime;

$count = $count+1;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('W'.$count,$totalOnTime)
        -> getStyle('W'.$count)
        -> applyFromArray($tableHeaderStyleFontBlackBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('W'.$count)
        -> getFill()
        -> setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        -> getStartColor()
        -> setRGB('f2f2f2');

$totalLate = $totalNCRLate+$totalProvLate;

$excel -> setActiveSheetIndex(0) 
        -> setCellValue('X'.$count,$totalLate)
        -> getStyle('X'.$count)
        -> applyFromArray($tableHeaderStyleFontRedBoldCalibri);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($centerText);

$excel -> setActiveSheetIndex(0) 
        -> getStyle('X'.$count)
        -> applyFromArray($border);

$excel -> setActiveSheetIndex(0) 
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
header('Content-Disposition: attachment; filename="DaysofDeliveryReport'.$sender.''.$deltype.''.$month.''.$year.'.xlsx"');
header('Cache-Control: max-age=0');
$file = PHPExcel_IOFactory::createWriter($excel,'Excel2007');
$file -> save('php://output');

mysqli_close($db_conn);

?>