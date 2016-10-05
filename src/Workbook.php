<?php

namespace ArneGroskurth\PHPExcelExtended;

use ArneGroskurth\TempFile\TempFile;


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
     * @var bool
     */
    protected static $setUp = false;

    /**
     * @var bool
     */
    protected static $isPdfAvailable = false;


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
    public function applyDefaultStyle(array $styles) {

        $this->phpExcel->getDefaultStyle()->applyFromArray($styles);

        return $this;
    }


    /**
     * @param bool $applyDefaultStyle
     *
     * @throws \PHPExcel_Exception
     */
    public function __construct($applyDefaultStyle = true) {

        if(!static::$setUp) {

            static::$setUp = true;

            // setup cache
            \PHPExcel_Settings::setCacheStorageMethod(\PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp, array(
                'memoryCacheSize' => '512MB'
            ));

            // setup pdf export
            if(class_exists('\TCPDF')) {

                $reflection = new \ReflectionClass('\TCPDF');

                if(!\PHPExcel_Settings::setPdfRenderer(\PHPExcel_Settings::PDF_RENDERER_TCPDF, dirname($reflection->getFileName()))) {

                    throw new \PHPExcel_Exception('Error setting up TCPDF as pdf rendering library!');
                }

                static::$isPdfAvailable = true;
            }
        }


        $this->phpExcel = new \PHPExcel();
        $this->phpExcel->removeSheetByIndex(0);

        if($applyDefaultStyle) {

            $this->applyDefaultStyle(array(
                'font' => array(
                    'name' => 'Calibri',
                    'size' => 10
                )
            ));
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

        return $this->getSheet($title);
    }


    /**
     * Renders the workbook as excel file and returns it as a temporary file.
     *
     * @return TempFile
     * @throws \PHPExcel_Exception
     */
    public function writeToTempFile() {

        try {

            $tempFile = new TempFile();
            $tempFile->accessPath(function($path) {

                $phpExcelWriter = new \PHPExcel_Writer_Excel2007($this->phpExcel);
                $phpExcelWriter->setPreCalculateFormulas(true);
                $phpExcelWriter->save($path);
            });

            return $tempFile;
        }
        catch(\Exception $e) {

            throw new \PHPExcel_Exception('Could not build response.', 0, $e);
        }
    }


    /**
     * @param string $fileName
     *
     * @return \Symfony\Component\HttpFoundation\Response
     * @throws \PHPExcel_Exception
     */
    public function buildResponse($fileName = 'Export.xlsx') {

        return $this->writeToTempFile()->buildResponse($fileName, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    }


    /**
     * Renders the workbook as pdf file and returns it as a temporary file.
     *
     * @param int $paperSize
     * @param string $orientation
     *
     * @return TempFile
     * @throws \PHPExcel_Exception
     */
    public function writePdfToTempFile($paperSize = \PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4, $orientation = \PHPExcel_Worksheet_PageSetup::ORIENTATION_PORTRAIT) {

        if(!self::$isPdfAvailable) {

            throw new \PHPExcel_Exception('No PDF writing library available. (Installing tecnickcom/tcpdf is suggested)');
        }


        // copy workbook and apply pageSetup settings
        if($paperSize !== null || $orientation !== null) {

            $workbook = $this->phpExcel->copy();

            foreach($workbook->getAllSheets() as $sheet) {

                $pageSetup = $sheet->getPageSetup();

                if($paperSize !== null) {

                    $pageSetup->setPaperSize($paperSize);
                }

                if($orientation !== null) {

                    $pageSetup->setOrientation($orientation);
                }
            }
        }

        else {

            $workbook = $this->phpExcel;
        }


        try {

            $tempFile = new TempFile();
            $tempFile->accessPath(function($path) use ($workbook) {

                $phpExcelWriter = new \PHPExcel_Writer_PDF($workbook);
                $phpExcelWriter->save($path);
            });

            return $tempFile;
        }
        catch(\Exception $e) {

            throw new \PHPExcel_Exception('Could not build response.', 0, $e);
        }
    }


    /**
     * @param string $fileName
     *
     * @return \Symfony\Component\HttpFoundation\Response
     * @throws \PHPExcel_Exception
     */
    public function buildPdfResponse($fileName = 'Export.pdf') {

        return $this->writePdfToTempFile()->buildResponse($fileName, 'application/pdf');
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