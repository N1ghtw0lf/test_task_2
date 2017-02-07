<?php
/**
 * Parse https://apps.skypeassets.com/offers/credit/rates?_accept=1.0&currency=USD&destination=IN&language=ru&origin=UA&seq=18%20Request%20Method:GET
 * and make xsl
 * Time: 16:40
 */

include('PHPExcel.php');
include('simple_html_dom.php');

$curl = curl_init();
curl_setopt($curl, CURLOPT_URL, "https://apps.skypeassets.com/offers/credit/rates?_accept=1.0&currency=USD&destination=IN&language=ru&origin=UA&seq=18%20Request%20Method:GET");
curl_setopt($curl, CURLOPT_RETURNTRANSFER, 1);
$output = curl_exec($curl);
curl_close($curl);

$infoCountry = [];
$infoMobile = [];
$infoFavourite = [];
$infoSms = [];

$html = new simple_html_dom();
$html->load($output);
$calls = $html->find('div.pstn-rates div.row');
foreach($calls as $call) {
    if ($call->class != 'row heading') {
        $title = trim($call->find('div.column p', 0)->plaintext);
        if (!strpos($title, ' – ')) {
            $infoCountry[] = [$title, trim($call->find('div.column p', 1)->plaintext)];
        } elseif (strpos($title, 'Мобильный')) {
            $infoMobile[] = [substr($title, 0, strpos($title, ' – ')), trim($call->find('div.column p', 1)->plaintext)];
        } else {
            $infoFavourite[] = [$title, trim($call->find('div.column p', 1)->plaintext)];
        }
    }
}
$sms = $html->find('div.sms-rates div.row');
foreach($sms as $message) {
    if ($message->class != 'row heading') {
        $infoSms[] = [trim($message->find('div.column p', 0)->plaintext), trim($message->find('div.column p', 1)->plaintext)];
    }
}

$xls = new PHPExcel();
$xls->removeSheetByIndex();
if (!empty($infoCountry)) {
    $workSheet = new PHPExcel_Worksheet($xls, 'Country');
    $xls->addSheet($workSheet);
    $xls->setActiveSheetIndexByName('Country');
    $xls->getActiveSheet()->fromArray($infoCountry, null, 'A1');
}
if (!empty($infoMobile)) {
    $workSheet = new PHPExcel_Worksheet($xls, 'Mobile');
    $xls->addSheet($workSheet);
    $xls->setActiveSheetIndexByName('Mobile');
    $xls->getActiveSheet()->fromArray($infoMobile, null, 'A1');
}
if (!empty($infoFavourite)) {
    $workSheet = new PHPExcel_Worksheet($xls, 'Favourite');
    $xls->addSheet($workSheet);
    $xls->setActiveSheetIndexByName('Favourite');
    $xls->getActiveSheet()->fromArray($infoFavourite, null, 'A1');
}
if (!empty($infoSms)) {
    $workSheet = new PHPExcel_Worksheet($xls, 'SMS');
    $xls->addSheet($workSheet);
    $xls->setActiveSheetIndexByName('SMS');
    $xls->getActiveSheet()->fromArray($infoSms, null, 'A1');
}
$xlsWriter = PHPExcel_IOFactory::createWriter($xls, 'Excel2007');
$xlsWriter->save('task.xls');
echo 'Creating xls successfully finished';