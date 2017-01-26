<?php

namespace ArneGroskurth\PHPExcelExtended;

class CellIterator extends \ArrayIterator
{
    use CoordinateMath;

    /**
     * @param string $coordinates
     */
    public function __construct($coordinates)
    {
        $cells = array();

        $origin = $this->getCoordinatesOrigin($coordinates);
        $rangeWidth = $this->getCoordinatesRangeWidth($coordinates);
        $rangeHeight = $this->getCoordinatesRangeHeight($coordinates);

        for ($row = 0; $row < ($rangeHeight - 1); $row++)
        {
            for ($column = 0; $column < ($rangeWidth - 1); $column++)
            {
                $cells[] = $this->addToCoordinates($origin, $column, $row);
            }
        }

        parent::__construct($cells);
    }
}
