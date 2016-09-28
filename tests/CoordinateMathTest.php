<?php

namespace ArneGroskurth\PHPExcelExtended\Tests;


use ArneGroskurth\PHPExcelExtended\CoordinateMath;

class CoordinateMathTest extends \PHPUnit_Framework_TestCase {

    use CoordinateMath;


    public function testGetCoordinatesRangeParts() {

        static::assertEquals(array('A1', 'B2'), $this->getCoordinatesRangeParts('A1:B2'));
    }


    /**
     * @depends testGetCoordinatesRangeParts
     */
    public function testGetCoordinatesOrigin() {

        static::assertEquals('B2', $this->getCoordinatesOrigin('B2'));
        static::assertEquals('B2', $this->getCoordinatesOrigin('B2:C2'));
        static::assertEquals('B2', $this->getCoordinatesOrigin('B2:B3'));
        static::assertEquals('B2', $this->getCoordinatesOrigin('B2:C3'));
        static::assertEquals('B2', $this->getCoordinatesOrigin('B2:ABC900'));
        static::assertEquals('AA100', $this->getCoordinatesOrigin('AA100:ABC900'));
    }


    public function testGetCoordinatesParts() {

        static::assertEquals(array('AB', '100'), $this->getCoordinatesParts('AB100'));
    }


    /**
     * @depends testGetCoordinatesParts
     */
    public function testGetCoordinatesColumnName() {

        static::assertEquals('X', $this->getCoordinatesColumnName('X1'));
        static::assertEquals('ABC', $this->getCoordinatesColumnName('ABC4231'));
    }


    /**
     * @depends testGetCoordinatesParts
     */
    public function testGetCoordinatesRowNumber() {

        static::assertEquals('9', $this->getCoordinatesRowNumber('C9'));
        static::assertEquals('4231', $this->getCoordinatesRowNumber('ABC4231'));
    }


    /**
     * @depends testGetCoordinatesRangeParts
     * @depends testGetCoordinatesRowNumber
     */
    public function testGetCoordinatesRangeHeight() {

        static::assertEquals(1, $this->getCoordinatesRangeHeight('B2:C2'));
        static::assertEquals(2, $this->getCoordinatesRangeHeight('B2:C3'));
        static::assertEquals(20, $this->getCoordinatesRangeHeight('B2:C21'));
    }


    public function testColumnNumberToColumnName() {

        static::assertEquals('C', $this->columnNumberToColumnName(2));
        static::assertEquals('AA', $this->columnNumberToColumnName(26));
    }


    public function testColumnNameToColumnNumber() {

        static::assertEquals('0', $this->columnNameToColumnNumber('A'));
        static::assertEquals('26', $this->columnNameToColumnNumber('AA'));
    }


    /**
     * @depends testGetCoordinatesRangeParts
     * @depends testGetCoordinatesColumnName
     * @depends testColumnNameToColumnNumber
     */
    public function testGetCoordinatesRangeWidth() {

        static::assertEquals(1, $this->getCoordinatesRangeWidth('A5:A10'));
        static::assertEquals(2, $this->getCoordinatesRangeWidth('A5:B5'));
        static::assertEquals(3, $this->getCoordinatesRangeWidth('A5:C5'));
        static::assertEquals(3, $this->getCoordinatesRangeWidth('A5:C8'));
        static::assertEquals(2, $this->getCoordinatesRangeWidth('B5:C8'));
    }


    /**
     * @depends testGetCoordinatesRangeParts
     * @depends testGetCoordinatesRowNumber
     * @depends testGetCoordinatesColumnName
     * @depends testColumnNumberToColumnName
     * @depends testColumnNameToColumnNumber
     */
    public function testAddToCoordinates() {

        static::assertEquals('AB101', $this->addToCoordinates('AA100', 1, 1));
        static::assertEquals('AA100', $this->addToCoordinates('AB101', -1, -1));
    }


    /**
     * @depends testAddToCoordinates
     */
    public function testAddToCoordinatesRef() {

        $coordinates = 'B3';

        $this->addToCoordinatesRef($coordinates, 1, 1);

        static::assertEquals('C4', $coordinates);
    }


    /**
     * @depends testAddToCoordinatesRef
     */
    public function testAddToCoordinatesReefAfter() {

        $coordinates = 'AA100';

        static::assertEquals('AA100', $this->addToCoordinatesRefAfter($coordinates, 1, 1));
        static::assertEquals('AB101', $coordinates);
    }


    /**
     * @depends testGetCoordinatesParts
     */
    public function testModifyCoordinates() {

        static::assertEquals('X100', $this->modifyCoordinates('A100', 'X', null));
        static::assertEquals('A1', $this->modifyCoordinates('A100', null, '1'));
    }
}