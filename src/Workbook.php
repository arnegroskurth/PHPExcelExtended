<?php

namespace ArneGroskurth\PHPExcelExtended;

use Symfony\Component\HttpFoundation\Response;


class Workbook {

    /**
     * @var \PHPExcel
     */
    protected $phpExcel;

    /**
     * @var array
     */
    protected $defaultStyle = array(
        'font' => array(
            'name' => 'Calibri',
            'size' => 10
        )
    );


    /**
     * @return \PHPExcel
     */
    public function getPhpExcel() {

        return $this->phpExcel;
    }


    /**
     * @param array $styles
     *
     * @return $this
     */
    public function setDefaultStyle(array $styles) {

        $this->defaultStyle = $styles;

        return $this;
    }


    /**
     * @return array
     */
    public function getDefaultStyle() {

        return $this->defaultStyle;
    }


    /**
     * @param \PHPExcel $PHPExcel
     *
     * @throws \PHPExcel_Exception
     */
    public function __construct(\PHPExcel $PHPExcel = null) {

        static $setUp = false;


        $this->phpExcel = $PHPExcel;

        if($this->phpExcel === null) {

            $this->phpExcel = new \PHPExcel();
            $this->phpExcel->removeSheetByIndex(0);
            $this->phpExcel->getDefaultStyle()->applyFromArray($this->defaultStyle);
        }


        if(!$setUp) {

            // setup cache
            $cacheMethod = \PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
            $cacheSettings = array('memoryCacheSize' => '512MB');
            \PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);


            // setup pdf export
            if(!\PHPExcel_Settings::setPdfRenderer(\PHPExcel_Settings::PDF_RENDERER_TCPDF, __DIR__ . '/../../../../vendor/tecnickcom/tcpdf')) {

                throw new \PHPExcel_Exception('Could not initialize PHPExcel PDF writer!');
            }


            $setUp = true;
        }
    }


    /**
     * @param string $title
     *
     * @return Sheet
     */
    public function getSheet($title) {

        if($sheet = $this->phpExcel->getSheetByName($title)) {

            return new Sheet($this, $sheet);
        }

        return null;
    }


    /**
     * @param string $title
     *
     * @return Sheet
     * @throws \PHPExcel_Exception
     */
    public function createSheet($title) {

        $sheet = $this->phpExcel->addSheet(new \PHPExcel_Worksheet($this->phpExcel));

        $sheet->setTitle($title);
        $sheet->getSheetView()->setZoomScale(80);

        return $this->getSheet($title);
    }


    /**
     * @param string $fileName
     *
     * @return Response
     * @throws \PHPExcel_Exception
     */
    public function buildResponse($fileName = 'Export') {

        $tmpFileName = tempnam('/tmp', 'WorkbookExport_');

        if($tmpFileName === false) {
            throw new \PHPExcel_Exception('Could not create temp file.');
        }

        $phpExcelWriter = new \PHPExcel_Writer_Excel2007($this->phpExcel);
        $phpExcelWriter->setPreCalculateFormulas(true);
        $phpExcelWriter->save($tmpFileName);

        $fileSize = filesize($tmpFileName);


        try {

            // build response
            $response = new Response(file_get_contents($tmpFileName));
            $response->headers->set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            $response->headers->set('Content-Disposition', sprintf('attachment; filename=\'%s.xlsx\'', $fileName));
            $response->headers->set('Content-Length', $fileSize);
        }
        catch(\Exception $e) {

            throw new \PHPExcel_Exception('Could not build response.', 0, $e);
        }

        unlink($tmpFileName);

        return $response;
    }


    /**
     * @param string $fileName
     *
     * @return Response
     * @throws \PHPExcel_Exception
     */
    public function buildPdfResponse($fileName = 'Export') {

        $tmpFileName = tempnam('/tmp', 'WorkbookExport_');

        if($tmpFileName === false) {
            throw new \PHPExcel_Exception('Could not create temp file.');
        }


        $this->phpExcel->getSheet(0)->getPageSetup()
            ->setOrientation(\PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE)
            ->setPaperSize(\PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4)
            ->setFitToPage()
        ;

        $phpExcelWriter = new \PHPExcel_Writer_PDF($this->phpExcel);
        $phpExcelWriter->save($tmpFileName);

        $fileSize = filesize($tmpFileName);


        // build response
        try {

            $response = new Response(file_get_contents($tmpFileName));
            $response->headers->set('Content-Type', 'application/pdf');
            $response->headers->set('Content-Disposition', sprintf('attachment; filename=\'%s.pdf\'', $fileName));
            $response->headers->set('Content-Length', $fileSize);
        }
        catch(\Exception $e) {

            throw new \PHPExcel_Exception('Could not build response.', 0, $e);
        }

        unlink($tmpFileName);

        return $response;
    }


    /**
     * @param array \PHPExcel_Style[]
     *
     * @return \PHPExcel_Style_Conditional[]
     */
    public static function createConditionalStyles(array $styles) {

        $return = array();

        foreach($styles as $value => $style) {

            $conditionalStyle = new \PHPExcel_Style_Conditional();
            $conditionalStyle->setConditionType(\PHPExcel_Style_Conditional::CONDITION_CELLIS);
            $conditionalStyle->setOperatorType(\PHPExcel_Style_Conditional::OPERATOR_EQUAL);
            $conditionalStyle->setConditions(array($value));
            $conditionalStyle->setStyle($style);

            $return[] = $conditionalStyle;
        }

        return $return;
    }


    /**
     * @param string $coordinates
     *
     * @return string
     */
    public static function getCoordinatesOrigin($coordinates) {

        $parts = static::getCoordinatesRangeParts($coordinates);

        return is_array($parts) ? $parts[0] : $coordinates;
    }


    /**
     * @param string $coordinates
     *
     * @return array|string
     */
    public static function getCoordinatesRangeParts($coordinates) {

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
    public static function getCoordinatesRangeWidth($coordinates) {

        $parts = static::getCoordinatesRangeParts($coordinates);

        return is_array($parts) ? static::alphaToColumnNumber(static::getCoordinatesColumn($parts[1])) - static::alphaToColumnNumber(static::getCoordinatesColumn($parts[0])) : 0;
    }


    /**
     * @param string $coordinates
     *
     * @return int
     * @throws \PHPExcel_Exception
     */
    public static function getCoordinatesRangeHeight($coordinates) {

        $parts = static::getCoordinatesRangeParts($coordinates);

        return is_array($parts) ? static::getCoordinatesRow($parts[1]) - static::getCoordinatesRow($parts[0]) : 0;
    }


    /**
     * @param string $coordinates
     * @param int $columns
     * @param int $rows
     *
     * @return string
     * @throws \PHPExcel_Exception
     */
    public static function addToCoordinates($coordinates, $columns = 0, $rows = 0) {

        $parts = static::getCoordinatesRangeParts($coordinates);

        if(is_array($parts)) {

            return sprintf('%s:%s', static::addToCoordinates($parts[0], $columns, $rows), static::addToCoordinates($parts[1], $columns, $rows));
        }

        else {

            $column = static::columnNumberToAlpha(static::alphaToColumnNumber(static::getCoordinatesColumn($coordinates)) + $columns);
            $row = static::getCoordinatesRow($coordinates) + $rows;

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
    public static function addToCoordinatesRef(&$coordinates, $columns = 0, $rows = 0) {

        return $coordinates = static::addToCoordinates($coordinates, $columns, $rows);
    }


    /**
     * @param string $coordinates
     * @param int $columns
     * @param int $rows
     *
     * @return string
     * @throws \PHPExcel_Exception
     */
    public static function addToCoordinatesRefAfter(&$coordinates, $columns = 0, $rows = 0) {

        $return = $coordinates;

        static::addToCoordinatesRef($coordinates, $columns, $rows);

        return $return;
    }


    /**
     * @param string $coordinates
     * @param string $column
     * @param int $row
     * @return string
     * @throws \PHPExcel_Exception
     */
    public static function modifyCoordinates($coordinates, $column = null, $row = null) {

        $parts = static::getCoordinatesParts($coordinates);

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
    public static function getCoordinatesParts($coordinates) {

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
    public static function getCoordinatesColumn($coordinates) {

        return static::getCoordinatesParts($coordinates)[0];
    }


    /**
     * @param $coordinates
     *
     * @return int
     * @throws \PHPExcel_Exception
     */
    public static function getCoordinatesRow($coordinates) {

        return static::getCoordinatesParts($coordinates)[1];
    }


    /**
     * http://stackoverflow.com/questions/3302857/algorithm-to-get-the-excel-like-column-name-of-a-number
     *
     * @param int $n
     *
     * @return string
     */
    public static function columnNumberToAlpha($n) {

        for($return = ''; $n >= 0; $n = (int)($n/26) - 1) {

            $return = chr($n % 26 + 0x41) . $return;
        }

        return $return;
    }


    /**
     * @param string $alpha
     *
     * @return int
     */
    public static function alphaToColumnNumber($alpha) {

        $return = 0;

        foreach(array_reverse(str_split($alpha)) as $index => $char) {

            $return += (ord($char) - 0x40) * pow(26, $index);
        }

        return $return - 1;
    }
}