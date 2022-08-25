<?php


ini_set('error_reporting', E_ALL);
ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);


require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Reader\Xlsx as ReaderXlsx;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as WriteXlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet as Worksheet;


$new_file = '../upload_files/'.$_FILES['file']['name'];

copy($_FILES['file']['tmp_name'], $new_file);



$inputFileName = $new_file;
$reader = new ReaderXlsx();
$spreadsheet = $reader->load($inputFileName);
$sheet = $spreadsheet->getActiveSheet();
$startWork = 9;

$worksheetInfo = $reader->listWorksheetInfo($inputFileName);
$totalRows = $worksheetInfo[0]['totalRows'];

$arr = [];
$totalV = [];

for ($row = 2; $row <= $totalRows; $row++) {
    $number = $sheet->getCell("A{$row}")->getValue();
    $name = $sheet->getCell("B{$row}")->getValue();
    $h = $sheet->getCell("C{$row}")->getValue();
    $vh = $sheet->getCell("D{$row}")->getValue();
    $v = $h * $vh;
    $totalV [] = $v;
    $startWeek = $sheet->getCell("F{$row}")->getValue();

    $arr [] = [
        'number' => $number,
        'name' => $name,
        'h' => $h,
        'vh' => $vh,
        'v' => $v,
        'startWeek' => $startWeek,
    ];

}

$extra = (array_sum($totalV) - 300) / 8;
$weekNorm = [];
foreach ($totalV as $item) {
    if ($item >= 13) {
        $weekNorm [] = $item - $extra;
    } else {
        $weekNorm [] = $item;
    }
}
$fin = [];
for ($i = 0; $i < count($arr); $i++) {
    $fin [] = [
        'number' => $arr[$i]['number'],
        'name' => $arr[$i]['name'],
        'h' => $arr[$i]['h'],
        'vh' => $arr[$i]['vh'],
        'v' => $arr[$i]['v'],
        'startWeek' => $arr[$i]['startWeek'],
        'workHour' => ceil($weekNorm[$i] / $arr[$i]['vh'])
    ];
}

$char = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];


for ($i = 1; $i <= 5; $i++){
    $myWorkSheet = new Worksheet($spreadsheet, 'new' . $i);
    $spreadsheet->addSheet($myWorkSheet);
    $writer = new WriteXlsx($spreadsheet);
    $sheet = $spreadsheet->getSheet($i);

    for($j = 1; $j<=count($fin)-1; $j++){
        $sheet->setCellValue($char[$j]. 1, $fin[$j]['name']);
        $sheet->setCellValue($char[$j] . 2, $fin[$j]['workHour']);
    }

}



$writer->save($inputFileName);

require_once '../views/vaucher.html';