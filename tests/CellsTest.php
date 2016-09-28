<?php

namespace ArneGroskurth\PHPExcelExtended\Tests;

use ArneGroskurth\PHPExcelExtended\Sheet;
use ArneGroskurth\PHPExcelExtended\Workbook;


class CellsTest extends \PHPUnit_Framework_TestCase {

    /**
     * @var Workbook
     */
    protected $workbook;

    /**
     * @var Sheet
     */
    protected $sheet;


    public function setUp() {

        $this->workbook = new Workbook();
        $this->sheet = $this->workbook->createSheet('Test');
    }


    public function testWriteRead() {

        $testContent = md5(time());

        $this->sheet->getCells('B3')->setValue($testContent);

        static::assertEquals($testContent, $this->sheet->getCells('B3')->getValue());
    }


    public function tearDown() {

        unset($this->sheet);
        unset($this->workbook);
    }
}