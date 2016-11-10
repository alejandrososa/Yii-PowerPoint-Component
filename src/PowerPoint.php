<?php

/**
 * Creado con PhpStorm.
 * Copyright (c) html 2016.
 * PowerPoint
 * Autor: Franklyn Alejandro Sosa Pérez <alesjohnson@hotmail.com>
 * Fecha: 24/08/2016
 * Hora: 10:05
 */

namespace AlejandroSosa\YiiPowerPoint;

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
//use PhpOffice\PhpPresentation\Slide\SlideMaster;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Slide\Background\Image;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Fill;

use AlejandroSosa\YiiPowerPoint\Common\ConstantesPPT;
use AlejandroSosa\YiiPowerPoint\ObjectsPptFactory;
use AlejandroSosa\YiiPowerPoint\Common\Helper;
use AlejandroSosa\YiiPowerPoint\Common\Style;


/**
 * Class PowerPoint
 * @package AlejandroSosa\YiiPowerPoint
 */
class PowerPoint extends \CApplicationComponent implements ConstantesPPT
{
    /**
     * Configuration options
     * @var array
     */
    public $options             = array();

    /**
     * Setting slides
     * @var array
     */
    public $slides              = array();

    /**
     * Path where file is saved
     * @var
     */
    private $_pathDir;

    /**
     * Default File name
     * @var string
     */
    private $_fileName          = 'presentation';

    /**
     * Default file extension
     * @var string
     */
    private $_fileExtension     = 'pptx';

    /**
     * Configuration file
     * @var array
     */
    private $_fileProperties    = array();

    /**
     * Path of file was saved
     * @var string
     */
    private $_filePath;

    /**
     * Configuration layout styles
     * @var array
     */
    private $_paramsLayout      = array();

    /**
     * @var PhpPresentation
     */
    private $_presentation;

    /**
     * @var SlideMaster
     */
    private $_masterSlide;

    /**
     * @var
     */
    private $_layoutSlide;

    /**
     * PowerPoint constructor.
     * PowerPoint Settings here
     */
    public function __construct()
    {
        // Create new PHPPresentation object
        $this->_presentation = new PhpPresentation();

        // fix need repair when open file
        $masterSlides        = $this->_presentation->getAllMasterSlides();
        $this->_masterSlide  = $masterSlides[0];
        $slides              = $this->_masterSlide->getAllSlideLayouts();
        $this->_layoutSlide  = $slides[0];
    }

    /**
     * Init all vars
     */
    public function init()
    {
        $options = $this->options;

        //file
        $this->_pathDir         = \Yii::app()->getBasePath() . '/runtime/ppt';
        $this->_fileName        = Helper::hasArrayProperty('fileName', $options) ? $options['fileName'] : $this->_fileName;
        $this->_fileExtension   = Helper::hasArrayProperty('fileExtension', $options) ? $options['fileExtension'] : $this->_fileExtension;

        //properties of file
        $this->_fileProperties  = Helper::hasArrayProperty('fileProperties', $options) ? $options['fileProperties'] : $this->_fileProperties;

        //layout of all slides
        $this->_paramsLayout    = Helper::hasArrayProperty('layout', $options) ? $options['layout'] : array();

        //path where the file is saved
        $this->_filePath        = $this->_pathDir .'/'. $this->_fileName .'.'. $this->_fileExtension;

        //directory for save file ppt
        $this->initStorage();
    }

    /**
     * Create presentation ppt
     * @param array $options
     * @param array $slides
     * @return string absolute path of file
     */
    public function generate($options = array(), $slides = array(), $download = true)
    {
        //init vars
        $this->options  = $options;
        $this->slides   = $slides;
        $this->init();

        //set properties informacion file
        $this->setPropertiesFile();

        //create slides
        $this->createCustomSlides();

        //save file ppt
        $file = $this->saveFile();

        //download file ppt
        if($download){
            $this->downloadFile();
        }

        return $file;
    }

    /**
     * Create custom slide with texts, images, tables and charts
     */
    private function createCustomSlides()
    {
        //Remove first slide
        $this->_presentation->removeSlideByIndex(0);

        foreach ($this->slides as $index => $slide) {
            //create slide
            $this->_presentation->createSlide();
            $this->_presentation->setActiveSlideIndex($index);

            //get current slide
            $current_slide = $this->_presentation->getActiveSlide();

            //set layout
            $current_slide->setSlideLayout($this->_layoutSlide);
            $this->assignBackground();

            //create and add objects to current slide
            foreach ($slide as $tipo => $options) {
                ObjectsPptFactory::build($tipo, $current_slide, $options);
            }
        }
    }

    //FILE

    /**
     * Save file PPT
     * The file is saved into runtime/ppt
     * @return string
     */
    private function saveFile()
    {
        if(!empty($this->_presentation)) {
            $oWriterPPTX = IOFactory::createWriter($this->_presentation, 'PowerPoint2007');
            $oWriterPPTX->save($this->_filePath);
            return $this->_filePath;
        }
    }

    /**
     * Download file generated ppt
     */
    private function downloadFile()
    {
        if(!empty($this->_filePath) && file_exists($this->_filePath)){
            header('Content-type:application/vnd.ms-powerpoint');
            header('Content-Disposition: attachment; filename="'.$this->_fileName.'.'.$this->_fileExtension.'"');
            header('Content-Length: ' . filesize($this->_filePath));
            readfile($this->_filePath);
            Yii::app()->end();
        }
    }

    /**
     * Check if storage directory exist or create it
     * The directory is created in runtime/
     */
    private function initStorage()
    {
        Helper::createDirectory($this->_pathDir);
    }

    /**
     * Set properties of file
     * Set the document information such as Title, Subject, Description, Creator, and Company name
     */
    private function setPropertiesFile()
    {
        if(!empty($this->options['fileProperties'])) {
            $properties = $this->options['fileProperties'];
            $creator    = Helper::hasArrayProperty('creator', $properties) ? $properties['creator'] : self::PPT_CREATOR;
            $title      = Helper::hasArrayProperty('title', $properties) ? $properties['title'] : self::PPT_TITLE;
            $subject    = Helper::hasArrayProperty('subject', $properties) ? $properties['subject'] : self::PPT_SUBJECT;
            $description= Helper::hasArrayProperty('description', $properties) ? $properties['description'] : self::PPT_DESCRIPTION;

            $this->_presentation->getDocumentProperties()
                ->setCreator($creator)
                ->setTitle($title)
                ->setSubject($subject)
                ->setDescription($description);
        }
    }


    //STYLE TEMPLATE

    /**
     * Assigns the background
     * @return bool
     */
    private function assignBackground()
    {
        if(empty($this->_paramsLayout) && empty($this->_paramsLayout['background'])){
            return false;
        }

        $current_slide = $this->_presentation->getActiveSlide();
        Style::setBackgroundSlide($current_slide, $this->_paramsLayout['background']);
    }

    //TODO class style function to move
    /**
     * Add logo into slide
     * @param PHPPresentation $objPHPPresentation
     * @return \PhpOffice\PhpPresentation\Slide
     */
    private function createLogo($objPHPPresentation)
    {
        // Create slide
        $slide = $objPHPPresentation->createSlide();

        // Add logo
        $shape = $slide->createDrawingShape();
        $shape->setName('PHPPresentation logo')
            ->setDescription('PHPPresentation logo')
            ->setPath(\Yii::getPathOfAlias('images') .'/ppt/logo.png')
            ->setHeight(36)
            ->setOffsetX(10)
            ->setOffsetY(10);
        //$shape->getShadow()->setVisible(true)->setDirection(45)->setDistance(10);

        // Return slide
        return $slide;
    }
}