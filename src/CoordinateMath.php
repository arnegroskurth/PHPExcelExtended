<?php

namespace ArneGroskurth\PHPExcelExtended;


trait CoordinateMath {

    /**
     * @param string $coordinates
     *
     * @return string
     */
    protected function getCoordinatesOrigin($coordinates) {

        $parts = $this->getCoordinatesRangeParts($coordinates);

        return is_array($parts) ? $parts[0] : $coordinates;
    }


    /**
     * @param string $coordinates
     *
     * @return array|string
     */
    protected function getCoordinatesRangeParts($coordinates) {

        if(preg_match('/^([A-Z]+[0-9]+):([A-Z]+[0-9]+)$/', $coordinates, $match)) {

            return array(
                $match[1],
                $match[2]
            );
        }

        return $coordinates;
    }


    /**
     * @param string $coordinates
     *
     * @return int
     * @throws \PHPExcel_Exception
     */
    protected function getCoordinatesRangeWidth($coordinates) {

        $parts = $this->getCoordinatesRangeParts($coordinates);

        return is_array($parts) ? $this->columnNameToColumnNumber($this->getCoordinatesColumnName($parts[1])) - $this->columnNameToColumnNumber($this->getCoordinatesColumnName($parts[0])) : 0;
    }


    /**
     * @param string $coordinates
     *
     * @return int
     * @throws \PHPExcel_Exception
     */
    protected function getCoordinatesRangeHeight($coordinates) {

        $parts = $this->getCoordinatesRangeParts($coordinates);

        return is_array($parts) ? $this->getCoordinatesRowNumber($parts[1]) - $this->getCoordinatesRowNumber($parts[0]) : 0;
    }


    /**
     * @param string $coordinates
     * @param int $columns
     * @param int $rows
     *
     * @return string
     * @throws \PHPExcel_Exception
     */
    protected function addToCoordinates($coordinates, $columns = 0, $rows = 0) {

        $parts = $this->getCoordinatesRangeParts($coordinates);

        if(is_array($parts)) {

            return sprintf('%s:%s', $this->addToCoordinates($parts[0], $columns, $rows), $this->addToCoordinates($parts[1], $columns, $rows));
        }

        else {

            $column = $this->columnNumberToColumnName($this->columnNameToColumnNumber($this->getCoordinatesColumnName($coordinates)) + $columns);
            $row = $this->getCoordinatesRowNumber($coordinates) + $rows;

            return sprintf('%s%s', $column, $row);
        }
    }


    /**
     * @param string $coordinates
     * @param int $columns
     * @param int $rows
     *
     * @return string
     * @throws \PHPExcel_Exception
     */
    protected function addToCoordinatesRef(&$coordinates, $columns = 0, $rows = 0) {

        return $coordinates = $this->addToCoordinates($coordinates, $columns, $rows);
    }


    /**
     * @param string $coordinates
     * @param int $columns
     * @param int $rows
     *
     * @return string
     * @throws \PHPExcel_Exception
     */
    protected function addToCoordinatesRefAfter(&$coordinates, $columns = 0, $rows = 0) {

        $return = $coordinates;

        $this->addToCoordinatesRef($coordinates, $columns, $rows);

        return $return;
    }


    /**
     * @param string $coordinates
     * @param string $column
     * @param int $row
     * @return string
     * @throws \PHPExcel_Exception
     */
    protected function modifyCoordinates($coordinates, $column = null, $row = null) {

        $parts = $this->getCoordinatesParts($coordinates);

        $parts[0] = $column ?: $parts[0];
        $parts[1] = $row ?: $parts[1];

        return sprintf('%s%d', $parts[0], $parts[1]);
    }


    /**
     * @param string $coordinates
     *
     * @return array
     * @throws \PHPExcel_Exception
     */
    protected function getCoordinatesParts($coordinates) {

        if(preg_match('/([A-Z]+)([0-9]+)/', $coordinates, $match)) {

            return array(
                $match[1],
                (int)$match[2]
            );
        }

        throw new \PHPExcel_Exception('Malformed coordinates.');
    }


    /**
     * @param $coordinates
     *
     * @return string
     * @throws \PHPExcel_Exception
     */
    protected function getCoordinatesColumnName($coordinates) {

        return $this->getCoordinatesParts($coordinates)[0];
    }


    /**
     * @param $coordinates
     *
     * @return int
     * @throws \PHPExcel_Exception
     */
    protected function getCoordinatesRowNumber($coordinates) {

        return $this->getCoordinatesParts($coordinates)[1];
    }


    /**
     * Converts base-10 column number to base-26 column name.
     *
     * @see http://stackoverflow.com/questions/3302857/algorithm-to-get-the-excel-like-column-name-of-a-number
     *
     * @param int $columnNumber
     *
     * @return string
     */
    protected function columnNumberToColumnName($columnNumber) {

        for($return = ''; $columnNumber >= 0; $columnNumber = (int)($columnNumber/26) - 1) {

            $return = chr($columnNumber % 26 + 0x41) . $return;
        }

        return $return;
    }


    /**
     * Converts base-26 column name to base-10 column number.
     *
     * @param string $columnName
     *
     * @return int
     */
    protected function columnNameToColumnNumber($columnName) {

        $return = 0;

        foreach(array_reverse(str_split($columnName)) as $index => $char) {

            $return += (ord($char) - 0x40) * pow(26, $index);
        }

        return $return - 1;
    }
}