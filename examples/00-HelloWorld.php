<?php

if(!include(__DIR__ . '/../vendor/autoload.php')) {

    exit('Could not include autoloader. Run "composer install" to install dependencies and dump autoloader.');
}


use ArneGroskurth\PHPExcelExtended\Workbook;


$workbook = new Workbook();
$workbook->createSheet('My Sheet!')->getCells('B2')->setValue('Hello World');

$workbook->buildResponse()->send();
