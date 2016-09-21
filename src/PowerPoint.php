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

use AlejandroSosa\YiiPowerPoint\Common\Charts;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Slide\Background\Image;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Fill;
use AlejandroSosa\YiiPowerPoint\Common\Helper;
use AlejandroSosa\YiiPowerPoint\Common\Tables;
use AlejandroSosa\YiiPowerPoint\Common\Style;


/**
 * Class PowerPoint
 * @package AlejandroSosa\YiiPowerPoint
 */
class PowerPoint extends \CApplicationComponent
{

    public $options             = [];
    public $slides              = [];

    private $_logo;
    private $_orientacion;
    private $_pathDir;
    private $_fileName          = 'Informe';
    private $_fileExtension     = 'pptx';
    private $_fileProperties    = [];
    private $_paramsLayout      = [];

    /**
     * @var PhpPresentation
     */
    private $_presentation;

    /**
     * @var Drawing
     */
    private $_shape;

    /* Properties file */
    const PPT_CREATOR                           = 'PHPOffice';
    const PPT_TITLE                             = 'Sample Title';
    const PPT_SUBJECT                           = 'Sample Subject';
    const PPT_DESCRIPTION                       = 'Sample Description';

    /* Alignment styles */
    const TEXT_ALIGN_HORIZONTAL_GENERAL         = 'l';
    const TEXT_ALIGN_HORIZONTAL_LEFT            = 'l';
    const TEXT_ALIGN_HORIZONTAL_RIGHT           = 'r';
    const TEXT_ALIGN_HORIZONTAL_CENTER          = 'ctr';
    const TEXT_ALIGN_HORIZONTAL_JUSTIFY         = 'just';
    const TEXT_ALIGN_HORIZONTAL_DISTRIBUTED     = 'dist';
    const TEXT_ALIGN_VERTICAL_BASE              = 'base';
    const TEXT_ALIGN_VERTICAL_AUTO              = 'auto';
    const TEXT_ALIGN_VERTICAL_BOTTOM            = 'b';
    const TEXT_ALIGN_VERTICAL_TOP               = 't';
    const TEXT_ALIGN_VERTICAL_CENTER            = 'ctr';
    const TEXT_HEIGHT                           = 300;
    const TEXT_WIDTH                            = 600;
    const TEXT_OFFSET_X                         = 170;
    const TEXT_OFFSET_Y                         = 180;
    const TEXT_SIZE                             = 14;
    const TEXT_BREAK                            = '<br>';

    const DEFAULT_COLOR                         = '00000000';

    const DEFAULT_MARGIN_LEFT                   = 50;
    const DEFAULT_MARGIN_TOP                    = 100;
    const DEFAULT_SLIDE_WITH                    = 1040;
    const DEFAULT_SLIDE_HEIGTH                  = 720;

    /**
     * PowerPoint constructor.
     * PowerPoint Settings here
     */
    public function __construct()
    {
        // Create new PHPPresentation object
        $this->_presentation = new PhpPresentation();
    }

    /**
     * Init all vars
     */
    public function init()
    {
        //file
        $this->_pathDir         = \Yii::app()->getBasePath() . '/runtime/ppt';
        $this->_fileName        = Helper::hasArrayProperty('fileName', $this->options) ? $this->options['fileName'] : $this->_fileName;
        $this->_fileExtension   = Helper::hasArrayProperty('fileExtension', $this->options) ? $this->options['fileExtension'] : $this->_fileExtension;

        //properties of file
        $this->_fileProperties  = Helper::hasArrayProperty('fileProperties', $this->options) ? $this->options['fileProperties'] : $this->_fileProperties;

        //layout of all slides
        $this->_paramsLayout = Helper::hasArrayProperty('layout', $this->options) ? $this->options['layout'] : [];

        //directory for save file ppt
        $this->initStorage();
    }


    /**
     * Create presentation ppt
     * @param array $options
     * @param array $slides
     */
    public function generate($options = [], $slides = [])
    {
        $this->options = $options;
        $this->slides = $slides;

        $this->init();

        //set properties informacion file
        $this->setPropertiesFile();

        //create slides
        $this->createCustomSlides();

        //download file ppt
        $this->saveFile();
    }

    //FILE

    /**
     * Save file PPT
     * The file is saved into runtime/ppt
     */
    public function saveFile()
    {
        if(!empty($this->_presentation)) {
            $path = $this->_pathDir .'/'. $this->_fileName .'.'. $this->_fileExtension;
            $oWriterPPTX = IOFactory::createWriter($this->_presentation, 'PowerPoint2007');
            $oWriterPPTX->save($path);
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
            $creator = !empty($this->options['fileProperties']['creator'])
                ? $this->options['fileProperties']['creator'] : self::PPT_CREATOR;
            $title = !empty($this->options['fileProperties']['title'])
                ? $this->options['fileProperties']['title'] : self::PPT_TITLE;
            $subject = !empty($this->options['fileProperties']['subject'])
                ? $this->options['fileProperties']['subject'] : self::PPT_SUBJECT;
            $description = !empty($this->options['fileProperties']['description'])
                ? $this->options['fileProperties']['description'] : self::PPT_DESCRIPTION;

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
        if(empty($this->_paramsLayout)
            && empty($this->_paramsLayout['background'])
            && file_exists($this->_paramsLayout['background'])){
            return false;
        }

        $current_slide = $this->_presentation->getActiveSlide();
        Style::setBackgroundSlide($current_slide, $this->_paramsLayout['background']);
    }

    /**
     * Add logo into slide
     *
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
//        $shape->getShadow()->setVisible(true)
//            ->setDirection(45)
//            ->setDistance(10);

        // Return slide
        return $slide;
    }

    //OBJECTS TEXT, IMAGES, ETC

    /**
     * Create object text into slide
     * @param array $params
     */
    private function createText($params = [])
    {
        $text       = Helper::hasArrayProperty('text', $params) ? $params['text'] : '';
        $options    = Helper::hasArrayProperty('options', $params) ? $params['options'] : [];
        $height     = Helper::hasArrayProperty('height', $options) ? $options['height'] : self::TEXT_HEIGHT;
        $width      = Helper::hasArrayProperty('width', $options) ? $options['width'] : self::TEXT_WIDTH;
        $offset_x   = Helper::hasArrayProperty('ox', $options) ? $options['ox'] : self::TEXT_OFFSET_X;
        $offset_y   = Helper::hasArrayProperty('oy', $options) ? $options['oy'] : self::TEXT_OFFSET_Y;
        $align      = Helper::hasArrayProperty('align', $options) ? $options['align'] : self::TEXT_ALIGN_HORIZONTAL_CENTER;
        $bold       = Helper::hasArrayProperty('bold', $options) ? $options['bold'] : false;
        $color      = Helper::hasArrayProperty('color', $options) ? $options['color'] : self::DEFAULT_COLOR;
        $size       = Helper::hasArrayProperty('size', $options) ? $options['size'] : self::TEXT_SIZE;

        $current_slide = $this->_presentation->getActiveSlide();
        $shape = $current_slide->createRichTextShape();

        //set height, width and offset rich text
        $shape->setHeight($height)->setWidth($width)->setOffsetX($offset_x)->setOffsetY($offset_y);

        //set align of text
        $paragraph = $shape->getActiveParagraph();
        Style::setAlignText($paragraph, $align);

        //check if text has break line
        if(Helper::stringContains($text, self::TEXT_BREAK)){
            foreach (Helper::convertStringToArray($text, self::TEXT_BREAK) as $item) {
                //set text
                $current_text = $shape->createTextRun($item);

                //set style
                Style::setStyleText($current_text, $size, $bold, $color);

                //add breakline
                $shape->createBreak();
            }
        }else{
            //set text
            $current_text = $shape->createTextRun($text);

            //set style
            Style::setStyleText($current_text, $size, $bold, $color);
        }
    }

    //TODO end this method for images
    /**
     * Create custom image into slide
     * @param array $params
     */
    private function createImage($params = [])
    {
        $height     = Helper::hasArrayProperty('height', $params) ? $params['height'] : self::TEXT_HEIGHT;
        $width      = Helper::hasArrayProperty('width', $params) ? $params['width'] : self::TEXT_WIDTH;
        $offset_x   = Helper::hasArrayProperty('ox', $params) ? $params['ox'] : self::TEXT_OFFSET_X;
        $offset_y   = Helper::hasArrayProperty('oy', $params) ? $params['oy'] : self::TEXT_OFFSET_Y;
        $name       = Helper::hasArrayProperty('name', $params) ? $params['name'] : '';
        $description= Helper::hasArrayProperty('decription', $params) ? $params['name'] : '';

//        if()

        $current_slide = $this->_presentation->getActiveSlide();
//        $shape = $current_slide->createRichTextShape();

        $shape = new Drawing\File();
        $shape->setName($name)->setDescription($description);
    }

    /**
     * Create custom table into slide
     * @param array $params
     */
    private function createTable($params = [])
    {
        //table
        $height     = Helper::hasArrayProperty('height', $params) ? $params['height'] : self::TEXT_HEIGHT;
        $width      = Helper::hasArrayProperty('width', $params) ? $params['width'] : self::TEXT_WIDTH;
        $offset_x   = Helper::hasArrayProperty('ox', $params) ? $params['ox'] : self::TEXT_OFFSET_X;
        $offset_y   = Helper::hasArrayProperty('oy', $params) ? $params['oy'] : self::TEXT_OFFSET_Y;
        $row_header = Helper::hasArrayProperty('header', $params) ? $params['header'] : [];
        $rows       = Helper::hasArrayProperty('rows', $params) ? $params['rows'] : [];

        //header
        $header_columns     = Helper::hasArrayProperty('columns', $row_header) ? $row_header['columns'] : [];
        $header_style       = Helper::hasArrayProperty('style', $row_header) ? $row_header['style'] : [];
        $header_background  = Helper::hasArrayProperty('background', $header_style) ? $header_style['background'] : 'FFFFFFFF';
        $header_text_bold   = Helper::hasArrayProperty('bold', $header_style) ? $header_style['bold'] : false;
        $header_text_size   = Helper::hasArrayProperty('size', $header_style) ? $header_style['size'] : self::TEXT_SIZE;
        $header_text_color  = Helper::hasArrayProperty('color', $header_style) ? $header_style['color'] : self::DEFAULT_COLOR;
        $header_text_align  = Helper::hasArrayProperty('align', $header_style) ? $header_style['align'] : self::TEXT_ALIGN_HORIZONTAL_CENTER;
        $header_width       = Helper::hasArrayProperty('width', $header_style) ? $header_style['width'] : 100;
        $header_height      = Helper::hasArrayProperty('height', $header_style) ? $header_style['height'] : 20;

        $col_total  = count($header_columns) > 0 ? count($header_columns) : 0;

        //get current slide
        $current_slide = $this->_presentation->getActiveSlide();

        //create a table shape
        $shape = Tables::createTable($current_slide, $col_total, $height, $width, $offset_x, $offset_y);

        //add row header
        Tables::createRow($shape, $header_columns, $header_text_size, $header_text_bold, $header_text_color,
            $header_text_align, $header_background, $header_width, $header_height);

        //add the remaining rows
        foreach ($rows as $row) {
            $texts = $row['columns'];
            $style = $row['style'];
            Tables::createRow($shape, $texts, $style['size'], $style['bold'], $style['color'], $style['align'], $style['background']);
        }
    }

    /**
     * Create custom chart pie 3D
     * @param array $params
     */
    private function createPie3D($params = []){
        $title       = Helper::hasArrayProperty('title', $params) ? $params['title'] : '';
        $series_data = Helper::hasArrayProperty('series', $params) ? $params['series'] : [];
        $options     = Helper::hasArrayProperty('options', $params) ? $params['options'] : [];

        //get current slide
        $current_slide = $this->_presentation->getActiveSlide();

        //generate chart
        Charts::createPie3D($current_slide, $title, $series_data, $options);
    }

    //ASSESORS

    /**
     * Create custom slide with texts, images and tables
     */
    private function createCustomSlides()
    {
        //Remove first slide
        $this->_presentation->removeSlideByIndex(0);

        foreach ($this->slides as $index => $slide) {
            //create slide
            $this->_presentation->createSlide();
            $this->_presentation->setActiveSlideIndex($index);

            //set layout
            $this->assignBackground();

            //add text
            if(Helper::hasArrayProperty('texts', $slide)){
                if(Helper::isMultiArray($slide['texts'])){
                    foreach ($slide['texts'] as $item) {
                        $this->createText($item);
                    }
                }else{
                    $this->createText($slide['texts']);
                }
            }

            //add image
            if(Helper::hasArrayProperty('images', $slide)) {
                if (Helper::isMultiArray($slide['images'])) {
                    foreach ($slide['images'] as $item) {
                        $this->createImage($item);
                    }
                } else {
                    $this->createImage($slide['images']);
                }
            }

            //add table
            if(Helper::hasArrayProperty('tables', $slide)) {
                if (Helper::isMultiArray($slide['tables'])) {
                    foreach ($slide['tables'] as $item) {
                        $this->createTable($item);
                    }
                } else {
                    $this->createTable($slide['tables']);
                }
            }

            //add pie 3D
            if(Helper::hasArrayProperty('pie3D', $slide)) {
                if (Helper::isMultiArray($slide['pie3D'])) {
                    foreach ($slide['pie3D'] as $item) {
                        $this->createPie3D($item);
                    }
                } else {
                    $this->createPie3D($slide['pie3D']);
                }
            }
        }
    }
}