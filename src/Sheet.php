<?php

namespace ArneGroskurth\PHPExcelExtended;


/**
 * This class represents a sheet within a workbook.
 *
 * @package ArneGroskurth\PHPExcelExtended
 */
class Sheet {

    use CoordinateMath;
    

    /**
     * @var Workbook
     */
    protected $workbook;

    /**
     * @var \PHPExcel_Worksheet
     */
    protected $worksheet;

    
    /**
     * @return Workbook
     */
    public function getWorkbook() {

        return $this->workbook;
    }

    
    /**
     * @return \PHPExcel_Worksheet
     */
    public function getWorksheet() {

        return $this->worksheet;
    }


    /**
     * @param Workbook $workbook
     * @param \PHPExcel_Worksheet $worksheet
     */
    public function __construct(Workbook $workbook, \PHPExcel_Worksheet $worksheet) {

        $this->workbook = $workbook;
        $this->worksheet = $worksheet;
    }


    /**
     * @param int $zoomScale
     *
     * @return $this
     */
    public function setZoomScale($zoomScale) {

        $this->worksheet->getSheetView()->setZoomScale($zoomScale);

        return $this;
    }


    /**
     * Provides an Excel-like
     *
     * @param string $coordinates
     *
     * @return Cells
     */
    public function getCells($coordinates) {

        return new Cells($this, $coordinates);
    }


    /**
     * @param array $widths
     * @param string $firstColumn
     *
     * @return $this
     */
    public function setColumnWidths(array $widths, $firstColumn = 'A') {

        // sequential array indices are replaced by column names starting with firstColumn
        if(array_keys($widths) === range(0, count($widths) - 1)) {

            $widths = array_combine(range($firstColumn, $this->columnNumberToColumnName(count($widths) - 1)), $widths);
        }

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
     * @param float $width
     *
     * @return Sheet
     */
    public function setSameColumnWidths($from, $to, $width = -1.0) {

        return $this->setColumnWidths(
            array_combine(
                range($from, $to),
                array_fill(0, $this->getCoordinatesRangeWidth(sprintf('%s1:%s1', $from, $to)), $width)
            )
        );
    }


    /**
     * @param int $row
     * @param float $height
     *
     * @return $this
     */
    public function setRowHeight($row, $height) {

        $this->worksheet->getRowDimension($row)->setRowHeight($height);

        return $this;
    }


    /**
     * @param string $column
     * @param float|bool $width
     *
     * @return $this
     */
    public function setColumnWidth($column, $width) {

        if (is_bool($width))
        {
            $this->worksheet->getColumnDimension($column)->setAutoSize($width);
        }
        else
        {
            $this->worksheet->getColumnDimension($column)->setWidth($width);
        }

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
     * @throws \PHPExcel_Exception
     */
    public function setBackground($color = 'FFFFFFFF', $toCells = 'AZ100', $additionalColumns = 100, $additionalRows = 100) {

        $this->getCells(sprintf('A1:%s', $this->addToCoordinates($toCells, $additionalColumns, $additionalRows)))->applyStyle(array(
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
