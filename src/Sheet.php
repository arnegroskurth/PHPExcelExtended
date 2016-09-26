<?php

namespace ArneGroskurth\PHPExcelExtended;


class Sheet {

    /**
     * @var Workbook
     */
    protected $excelReporter;

    /**
     * @var \PHPExcel_Worksheet
     */
    protected $worksheet;

    
    /**
     * @return Workbook
     */
    public function getWorkbook() {

        return $this->excelReporter;
    }

    
    /**
     * @return \PHPExcel_Worksheet
     */
    public function getWorksheet() {

        return $this->worksheet;
    }


    /**
     * @param Workbook $excelReporter
     * @param \PHPExcel_Worksheet $worksheet
     */
    public function __construct(Workbook $excelReporter, \PHPExcel_Worksheet $worksheet) {

        $this->excelReporter = $excelReporter;
        $this->worksheet = $worksheet;
    }


    /**
     * @param string $coordinates
     *
     * @return Cells
     */
    public function getCells($coordinates) {

        return new Cells($this, $coordinates);
    }


    /**
     * @param array $widths
     *
     * @return $this
     */
    public function setColumnWidths(array $widths) {

        foreach($widths as $column => $width) {

            $columnDimension = $this->worksheet->getColumnDimension($column);

            if(is_int($width)) {
                $columnDimension->setWidth($width);
            }

            elseif(is_bool($width)) {
                $columnDimension->setAutoSize($width);
            }
        }

        return $this;
    }


    /**
     * @param string $from
     * @param string $to
     * @param int $width
     *
     * @return Sheet
     */
    public function setSameColumnWidths($from, $to, $width = -1) {

        return $this->setColumnWidths(
            array_combine(
                range($from, $to),
                array_fill(0, Workbook::getCoordinatesRangeWidth(sprintf('%s1:%s1', $from, $to)) + 1, $width)
            )
        );
    }


    /**
     * @param int $row
     * @param int $height
     *
     * @return $this
     */
    public function setRowHeight($row, $height) {

        $this->worksheet->getRowDimension($row)->setRowHeight($height);

        return $this;
    }


    /**
     * Todo: Remove toCells-argument and apply color to entire worksheet (how???)
     *
     * @param string $color
     * @param string $toCells
     * @param int $additionalColumns
     * @param int $additionalRows
     *
     * @return $this
     */
    public function setBackground($color = 'FFFFFFFF', $toCells = 'AZ100', $additionalColumns = 100, $additionalRows = 100) {

        $this->getCells(sprintf('A1:%s', Workbook::addToCoordinates($toCells, $additionalColumns, $additionalRows)))->applyStyle(array(
            'fill' => array(
                'type' => \PHPExcel_Style_Fill::FILL_SOLID,
                'startcolor' => array(
                    'argb' => $color
                )
            )
        ));

        return $this;
    }
}