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
    <!-- <script type="text/javascript">
        window.onload = function() { window.print(); }
    </script> -->

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
        .span-border-bottom {
            border: 1px solid #999;
            border-radius: 5px;
            width: 100%;
            height: 112px;
            padding: 4px 5px;
            margin: 10px 0px 10px 0px;
        }
        .rts-border {
            border: 1px solid #999;
            width: 100%;
            padding: 3px;
            margin: 10px 0px 5px 0px;
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
            #page-break {margin-top: 18mm;}
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

$count = 0; 
$recid = "1";
$cyclecode = "1231231";
$company = "sdfsdfsdf";
$subsname = "sdfsdfsdf";
$barcode = "dfgfsgdfgsdfg";
$area = "SDfgsdfgsdfgdfg";
$address = "sdfgsdfgsdfgsdf";
$deltype = "asdasddas";
$pud = "adfasdfadf";
$jobnumber = "aasdfsdfsfds";
$seqnumber = "dsfgsdfgsdfgs";

$newPud = "dfasdfasdf";
                       
                       for($i = 0; $i < 10; $i++) {
                           $count++;
                        if($count % 4 != 0) {

                            if($count % 4 == 1) {
    
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
                                    <tr class="highlight">
                                        <td colspan="4">
                                            <div class="span-border-bottom">
                                            <b style="font-size: 9px;">Paalala para sa tamang proseso ng pag deliver.</b><br>
                                            <i>Ang pagpeke ng pirma sa resibo ay may kaukulang mabigat na parusa.</i><br>
                                            1.Magtanong muna kung tama ang address na iyong dedeliveran at kung doon ba talaga nakatira ang subscriber.<br>
                                            2.Iparesib at papirmahan ang resibo.<br>
                                            3.Hindi mensahero ang pipirma maliban kung ayaw talaga pumirma ng magreresib ilagay ang pangalan ng nag resib relation at kulay ng gate.<br>
                                            4.Maari lamang mag sulat ng mailbox sa resibo kung talagang may mailbox talaga at tama ang dinedeliveran. Kung may mailbox ilagay ang kulay ng mailbox.
                                            </div>
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
                                    <tr class="highlight">
                                        <td colspan="4">
                                            <div class="span-border-bottom">
                                            <b style="font-size: 9px;">Paalala para sa tamang proseso ng pag deliver.</b><br>
                                            <i>Ang pagpeke ng pirma sa resibo ay may kaukulang mabigat na parusa.</i><br>
                                            1.Magtanong muna kung tama ang address na iyong dedeliveran at kung doon ba talaga nakatira ang subscriber.<br>
                                            2.Iparesib at papirmahan ang resibo.<br>
                                            3.Hindi mensahero ang pipirma maliban kung ayaw talaga pumirma ng magreresib ilagay ang pangalan ng nag resib relation at kulay ng gate.<br>
                                            4.Maari lamang mag sulat ng mailbox sa resibo kung talagang may mailbox talaga at tama ang dinedeliveran. Kung may mailbox ilagay ang kulay ng mailbox.
                                            </div>
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
                                    <tr class="highlight">
                                        <td colspan="4">
                                            <div class="span-border-bottom">
                                            <b style="font-size: 9px;">Paalala para sa tamang proseso ng pag deliver.</b><br>
                                            <i>Ang pagpeke ng pirma sa resibo ay may kaukulang mabigat na parusa.</i><br>
                                            1.Magtanong muna kung tama ang address na iyong dedeliveran at kung doon ba talaga nakatira ang subscriber.<br>
                                            2.Iparesib at papirmahan ang resibo.<br>
                                            3.Hindi mensahero ang pipirma maliban kung ayaw talaga pumirma ng magreresib ilagay ang pangalan ng nag resib relation at kulay ng gate.<br>
                                            4.Maari lamang mag sulat ng mailbox sa resibo kung talagang may mailbox talaga at tama ang dinedeliveran. Kung may mailbox ilagay ang kulay ng mailbox.
                                            </div>
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
                       
                       ?>
                            

    </div>

    <!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
    <!-- Include all compiled plugins (below), or include individual files as needed -->
    <script src="js/bootstrap.min.js"></script>
  </body>
</html>