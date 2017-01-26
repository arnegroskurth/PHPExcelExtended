<?php

namespace ArneGroskurth\PHPExcelExtended;

/**
 * Class CellIterator
 *
 * @package ArneGroskurth\PHPExcelExtended
 */
class CellIterator extends \ArrayIterator
{
    use CoordinateMath;

    const ITERATE_COLUMNS = 1;
    const ITERATE_ROWS = 2;
    const ITERATE_ALL = 3;

    /**
     * @param Cells $cells
     * @param int $mode
     */
    public function __construct(Cells $cells, $mode = self::ITERATE_ALL)
    {
        $array = array();

        $origin = $this->getCoordinatesOrigin($cells->getCoordinates());
        $rangeWidth = $this->getCoordinatesRangeWidth($cells->getCoordinates());
        $rangeHeight = $this->getCoordinatesRangeHeight($cells->getCoordinates());

        for ($row = 0; $row < $rangeHeight; $row++)
        {
            for ($column = 0; $column < $rangeWidth; $column++)
            {
                $array[] = $cells->getSheet()->getCells($this->addToCoordinates($origin, $column, $row));

                if ($column === 0 && !($mode & self::ITERATE_COLUMNS)) break;
            }

            if ($row === 0 && !($mode & self::ITERATE_ROWS)) break;
        }

        parent::__construct($array);
    }

    /**
     * @return Cells
     */
    public function current()
    {
        return parent::current();
    }
}
