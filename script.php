<?php

// CONNECT TO DATABASE
// ! CHANGE THESE SETTINGS TO YOUR OWN !
define('DB_HOST', 'localhost');
define('DB_NAME', 'digitalis_db');
define('DB_CHARSET', 'utf8');
define('DB_USER', 'root');
define('DB_PASSWORD', 'root@1234');
$pdo = new PDO(
    "mysql:host=".DB_HOST.";dbname=".DB_NAME.";charset=".DB_CHARSET,
    DB_USER, DB_PASSWORD, [
        PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
        PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
        PDO::ATTR_EMULATE_PREPARES => false,
    ]
);


require "vendor/autoload.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


// CREATE A NEW SPREADSHEET + SET METADATA
$spreadsheet = new Spreadsheet();
$spreadsheet->getProperties()
    ->setTitle('Product Price History');


// NEW WORKSHEET
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Price');

// POPULATING DATA
$stmt = $pdo->prepare("SELECT DATE_FORMAT(date, '%Y-%m-%d') as date,name,ean,sku,price from digitalis_crawl_history ORDER BY ean DESC");
$stmt->execute();

// STYLING EXCEL COLUMN AND ROW
$i = 1;
$sheet->getColumnDimension('B')->setAutoSize(true);
$sheet->getStyle('B:D')->getAlignment()->setHorizontal('right');
$sheet->getStyle('A1')->getAlignment()->setHorizontal('center');
$sheet->getStyle('E')->getAlignment()->setHorizontal('center');
$sheet->getStyle("A1:E1")->getFont()->setBold( true );

$sheet->setCellValue('A'.$i, 'PRODUCT NAME');
$sheet->setCellValue('B'.$i, 'EAN');
$sheet->setCellValue('C'.$i, 'SKU');
$sheet->setCellValue('D'.$i, 'PRICE');
$sheet->setCellValue('E'.$i, 'DATE');

$i = 2;
while ($row = $stmt->fetch(PDO::FETCH_NAMED)) {
    $sheet->setCellValue('A'.$i, $row['name']);
    $sheet->setCellValue('B'.$i, $row['ean']);
    $sheet->getStyle('B'.$i)
          ->getNumberFormat()
          ->setFormatCode('0');
    $sheet->setCellValue('C'.$i, $row['sku']);
    $sheet->setCellValue('D'.$i, $row['price']);
    $sheet->setCellValue('E'.$i, $row['date']);
    $i++;
}

// OUTPUT
$writer = new Xlsx($spreadsheet);

// DOWNLOADING THE SCRIPT
$ts = gmdate("D, d M Y H:i:s") . " GMT";
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="price_history.xlsx"');
header("Expires: $ts");
header("Last-Modified: $ts");
header("Pragma: no-cache");
header("Cache-Control: no-cache, must-revalidate");
header('Pragma: public');
$writer->save('php://output');
?>