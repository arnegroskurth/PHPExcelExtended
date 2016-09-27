<?php

namespace ArneGroskurth\PHPExcelExtended;

use ArneGroskurth\TempFile\TempFile;
use Symfony\Component\HttpFoundation\Response;


/**
 * This class represents a workbook and additionally provides common functions e.g. to convert between different coordinate formats.
 *
 * @package ArneGroskurth\PHPExcelExtended
 */
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
            /*if(!\PHPExcel_Settings::setPdfRenderer(\PHPExcel_Settings::PDF_RENDERER_TCPDF, __DIR__ . '/../../../../vendor/tecnickcom/tcpdf')) {

                throw new \PHPExcel_Exception('Could not initialize PHPExcel PDF writer!');
            }*/


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
    public function buildResponse($fileName = 'Export.xlsx') {

        try {

            $tempFile = new TempFile();
            $tempFile->accessPath(function($path) {

                $phpExcelWriter = new \PHPExcel_Writer_Excel2007($this->phpExcel);
                $phpExcelWriter->setPreCalculateFormulas(true);
                $phpExcelWriter->save($path);
            });

            return $tempFile->buildResponse($fileName, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

        }
        catch(\Exception $e) {

            throw new \PHPExcel_Exception('Could not build response.', 0, $e);
        }
    }


    /**
     * @param string $fileName
     *
     * @return Response
     * @throws \PHPExcel_Exception
     */
    public function buildPdfResponse($fileName = 'Export.pdf') {

        $this->phpExcel->getSheet(0)->getPageSetup()
            ->setOrientation(\PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE)
            ->setPaperSize(\PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4)
            ->setFitToPage()
        ;

        try {

            $tempFile = new TempFile();
            $tempFile->accessPath(function($path) {

                $phpExcelWriter = new \PHPExcel_Writer_PDF($this->phpExcel);
                $phpExcelWriter->save($path);
            });

            return $tempFile->buildResponse($fileName, 'application/pdf');

        }
        catch(\Exception $e) {

            throw new \PHPExcel_Exception('Could not build response.', 0, $e);
        }
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
}