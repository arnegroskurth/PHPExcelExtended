<?php

namespace ArneGroskurth\PHPExcelExtended;


/**
 * This class represents a cell selection on a given sheet and mainly provides styling and formatting functionality.
 *
 * @package ArneGroskurth\PHPExcelExtended
 */
class Cells {

    /**
     * @var Sheet
     */
    protected $sheet;

    /**
     * @var string
     */
    protected $coordinates;


    /**
     * @return Sheet
     */
    public function getSheet() {

        return $this->sheet;
    }


    /**
     * @return string
     */
    public function getCoordinates() {

        return $this->coordinates;
    }


    /**
     * @param Sheet $sheet
     * @param string $coordinates
     */
    public function __construct(Sheet $sheet, $coordinates) {

        $this->sheet = $sheet;
        $this->coordinates = $coordinates;
    }


    /**
     * @param mixed $value
     *
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function setValue($value) {

        if($value instanceof \DateTime) {

            $value = \PHPExcel_Shared_Date::PHPToExcel($value);
        }


        // set row of values
        if(is_array($value)) {

            $currentCoordinates = Workbook::getCoordinatesOrigin($this->coordinates);

            foreach($value as $val) {

                $this->getPHPExcelCell($currentCoordinates)->setValue($val);

                $currentCoordinates = Workbook::addToCoordinates($currentCoordinates, 1);
            }
        }

        // merge cells if range coordinates given and set single value
        else {

            if(Workbook::getCoordinatesRangeWidth($this->coordinates) > 0) {

                $this->sheet->getWorksheet()->mergeCells($this->coordinates);
            }

            $this->getPHPExcelCell(Workbook::getCoordinatesOrigin($this->coordinates))->setValue($value);
        }

        return $this;
    }


    /**
     * @param array $style
     *
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function applyStyle(array $style) {

        $this->getPHPExcelStyle()->applyFromArray($style);

        return $this;
    }


    /**
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function styleCentered() {

        $this->getPHPExcelStyle()->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

        return $this;
    }


    /**
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function styleBold() {

        $this->getPHPExcelStyle()->getFont()->setBold(true);

        return $this;
    }


    /**
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function styleItalic() {

        $this->getPHPExcelStyle()->getFont()->setItalic(true);

        return $this;
    }


    /**
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function styleUnderlined() {

        $this->getPHPExcelStyle()->getFont()->setUnderline(true);

        return $this;
    }


    /**
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function styleStrikethrough() {

        $this->getPHPExcelStyle()->getFont()->setStrikethrough(true);

        return $this;
    }


    /**
     * @param string $color
     * @param string $style
     *
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function styleWithBorder($color = 'FF000000', $style = \PHPExcel_Style_Border::BORDER_THIN) {

        $this->applyStyle(array(
            'borders' => array(
                'allborders' => array(
                    'style' => $style,
                    'color' => array(
                        'argb' => $color
                    )
                )
            )
        ));

        return $this;
    }


    /**
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function stylesAsFloat() {

        $this->getPHPExcelStyle()->applyFromArray(array(
            'numberformat' => array(
                'code' => '#,##0.00'
            )
        ));

        return $this;
    }


    /**
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function styleAsDate() {

        $this->getPHPExcelStyle()->applyFromArray(array(
            'numberformat' => array(
                'code' => 'dd.mm.yyyy'
            )
        ));

        return $this;
    }


    /**
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function styleAsMonthOfYear() {

        $this->getPHPExcelStyle()->applyFromArray(array(
            'numberformat' => array(
                'code' => 'mmmm yyyy'
            )
        ));

        return $this;
    }


    /**
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function styleAsCurrency() {

        $this->getPHPExcelStyle()->applyFromArray(array(
            'numberformat' => array(
                'code' => '#,##0.00 €'
            )
        ));

        return $this;
    }


    /**
     * @param int $height
     *
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function setRowHeight($height) {

        $rangeParts = Workbook::getCoordinatesRangeParts($this->coordinates);
        $row = Workbook::getCoordinatesRow(is_array($rangeParts) ? $rangeParts[0] : $rangeParts);

        foreach(range($row, $row + Workbook::getCoordinatesRangeHeight($this->coordinates)) as $row) {

            $this->sheet->setRowHeight($row, $height);
        }

        return $this;
    }


    /**
     * @param bool $wrapText
     *
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function setWrapText($wrapText) {

        $this->sheet->getWorksheet()->getStyle($this->coordinates)->getAlignment()->setWrapText($wrapText);

        return $this;
    }


    /**
     * @param string $coordinates
     *
     * @return \PHPExcel_Cell
     * @throws \PHPExcel_Exception
     */
    protected function getPHPExcelCell($coordinates = null) {

        return $this->sheet->getWorksheet()->getCell($coordinates ?: $this->coordinates);
    }


    /**
     * @param string $coordinates
     *
     * @return \PHPExcel_Style
     * @throws \PHPExcel_Exception
     */
    protected function getPHPExcelStyle($coordinates = null) {

        return $this->sheet->getWorksheet()->getStyle($coordinates ?: $this->coordinates);
    }
}