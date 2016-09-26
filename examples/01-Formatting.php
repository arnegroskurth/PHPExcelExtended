<?php

if(!include(__DIR__ . '/../vendor/autoload.php')) {

    exit('Could not include autoloader. Run "composer install" to install dependencies and dump autoloader.');
}


use ArneGroskurth\PHPExcelExtended\Workbook;


$workbook = new Workbook();
$sheet = $workbook->createSheet('My Sheet');

$sheet->getCells('B2')->setValue('Hello World!')->styleBold();
$sheet->getCells('B3')->setValue('Hello World!')->styleItalic();
$sheet->getCells('B4')->setValue('Hello World!')->styleUnderlined();
$sheet->getCells('B5')->setValue('Hello World!')->styleStrikethrough();
$sheet->getCells('B6')->setValue('Hello World!')->styleBold()->styleItalic()->styleUnderlined()->styleStrikethrough();

$workbook->buildResponse()->send();
