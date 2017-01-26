<?php

namespace ArneGroskurth\PHPExcelExtended;


/**
 * This class represents a cell selection on a given sheet and mainly provides styling and formatting functionality.
 *
 * @package ArneGroskurth\PHPExcelExtended
 */
class Cells implements \IteratorAggregate {
    
    use CoordinateMath;
    

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

            $this->formatAsDate();
        }


        // set row of values
        if(is_array($value)) {

            $currentCoordinates = $this->getCoordinatesOrigin($this->coordinates);

            foreach($value as $val) {

                $this->sheet->getCells($this->addToCoordinatesRefAfter($currentCoordinates, 1))->setValue($val);
            }
        }

        // merge cells if range coordinates given and set single value
        else {

            if($this->getCoordinatesRangeWidth($this->coordinates) > 1) {

                $this->sheet->getWorksheet()->mergeCells($this->coordinates);
            }

            $this->getPHPExcelCell($this->getCoordinatesOrigin($this->coordinates))->setValue($value);
        }

        return $this;
    }


    /**
     * @return mixed
     * @throws \PHPExcel_Exception
     */
    public function getValue() {

        return $this->getPHPExcelCell()->getValue();
    }


    /**
     * @param string $format
     *
     * @return $this
     */
    public function formatAsDate($format = 'dd.mm.yyyy') {

        $this->getPHPExcelStyle()->getNumberFormat()->setFormatCode($format);

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
    public function styleCenteredVertically() {

        $this->getPHPExcelStyle()->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);

        return $this;
    }


    /**
     * @param string $color
     * @param string $fillType
     *
     * @return $this
     */
    public function styleBackground($color = \PHPExcel_Style_Color::COLOR_WHITE, $fillType = \PHPExcel_Style_Fill::FILL_SOLID) {

        $fill = $this->getPHPExcelStyle()->getFill();

        $fill->setStartColor(new \PHPExcel_Style_Color($color));
        $fill->setFillType($fillType);

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

        foreach ($this->getIterator(CellIterator::ITERATE_ROWS) as $cell)
        {
            $this->sheet->setRowHeight($this->getCoordinatesRowNumber($cell->getCoordinates()), $height);
        }

        return $this;
    }


    /**
     * @param float $width
     *
     * @return $this
     */
    public function setColumnWidth($width) {

        foreach ($this->getIterator(CellIterator::ITERATE_COLUMNS) as $cell)
        {
            $this->sheet->setColumnWidth($this->getCoordinatesColumnName($cell->getCoordinates()), $width);
        }

        return $this;
    }


    /**
     * @param bool $wrapText
     *
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function setWrapText($wrapText = true) {

        $this->sheet->getWorksheet()->getStyle($this->coordinates)->getAlignment()->setWrapText($wrapText);

        return $this;
    }


    /**
     * @param int $value
     *
     * @return $this
     */
    public function setTextRotation($value)
    {
        $this->sheet->getWorksheet()->getStyle($this->coordinates)->getAlignment()->setTextRotation($value);

        return $this;
    }


    /**
     * @param int $mode
     *
     * @return CellIterator
     */
    public function getIterator($mode = CellIterator::ITERATE_ALL)
    {
        return new CellIterator($this, $mode);
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
