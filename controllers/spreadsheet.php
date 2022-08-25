<?php


require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Reader\Xlsx as ReaderXlsx;

$inputFileName = 'Drivers.xlsx';
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

require_once  'views/vaucher.html';



