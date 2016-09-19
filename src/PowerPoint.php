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
use PhpOffice\PhpPresentation\Style\Color;
//use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Slide\Background\Image;
use PhpOffice\PhpPresentation\Shape\RichText;

use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Fill;

use AlejandroSosa\YiiPowerPoint\Common\Helper;
use AlejandroSosa\YiiPowerPoint\Common\Alignment;

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

        $bkImage = new Image();
        $bkImage->setPath($this->_paramsLayout['background']);

        $current_slide = $this->_presentation->getActiveSlide();
        $current_slide->setBackground($bkImage);
    }

    /**
     * Creates a templated slide
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
        $height     = Helper::hasArrayProperty('height', $params) ? $params['height'] : self::TEXT_HEIGHT;
        $width      = Helper::hasArrayProperty('width', $params) ? $params['width'] : self::TEXT_WIDTH;
        $offset_x   = Helper::hasArrayProperty('ox', $params) ? $params['ox'] : self::TEXT_OFFSET_X;
        $offset_y   = Helper::hasArrayProperty('oy', $params) ? $params['oy'] : self::TEXT_OFFSET_Y;
        $align      = Helper::hasArrayProperty('align', $params) ? $params['align'] : self::TEXT_ALIGN_HORIZONTAL_CENTER;
        $text       = Helper::hasArrayProperty('text', $params) ? $params['text'] : '';
        $bold       = Helper::hasArrayProperty('bold', $params) ? $params['bold'] : false;
        $color      = Helper::hasArrayProperty('color', $params) ? $params['color'] : self::DEFAULT_COLOR;
        $size       = Helper::hasArrayProperty('size', $params) ? $params['size'] : self::TEXT_SIZE;

        $current_slide = $this->_presentation->getActiveSlide();
        $shape = $current_slide->createRichTextShape();

        //set height, width and offset rich text
        $shape->setHeight($height)->setWidth($width)->setOffsetX($offset_x)->setOffsetY($offset_y);

        //set align of text
        $paragraph = $shape->getActiveParagraph();
        Alignment::setAlignText($paragraph, $align);

        //check if text has break line
        if(Helper::stringContains($text, self::TEXT_BREAK)){
            foreach (Helper::convertStringToArray($text, self::TEXT_BREAK) as $item) {
                //set text
                $current_text = $shape->createTextRun($item);

                //set style
                $current_text->getFont()->setBold($bold);
                $current_text->getFont()->setSize($size);
                $current_text->getFont()->setColor( new Color($color) );

                //add breakline
                $shape->createBreak();
            }
        }else{
            //set text
            $current_text = $shape->createTextRun($text);

            //set style
            $current_text->getFont()->setBold($bold);
            $current_text->getFont()->setSize($size);
            $current_text->getFont()->setColor( new Color($color) );
        }
    }

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

    private function createTable($params = [])
    {
        //table
        $height     = Helper::hasArrayProperty('height', $params) ? $params['height'] : self::TEXT_HEIGHT;
        $width      = Helper::hasArrayProperty('width', $params) ? $params['width'] : self::TEXT_WIDTH;
        $offset_x   = Helper::hasArrayProperty('ox', $params) ? $params['ox'] : self::TEXT_OFFSET_X;
        $offset_y   = Helper::hasArrayProperty('oy', $params) ? $params['oy'] : self::TEXT_OFFSET_Y;
        $col_header = Helper::hasArrayProperty('header', $params) ? $params['header'] : [];

        //header
        $header_titles      = Helper::hasArrayProperty('titles', $col_header) ? $col_header['titles'] : [];
        $header_style       = Helper::hasArrayProperty('style', $col_header) ? $col_header['style'] : [];
        $header_background  = Helper::hasArrayProperty('background', $header_style) ? $header_style['background'] : '';
        $header_text_bold   = Helper::hasArrayProperty('bold', $header_style) ? $header_style['bold'] : '';
        $header_text_size   = Helper::hasArrayProperty('size', $header_style) ? $header_style['size'] : '';
        $header_text_color  = Helper::hasArrayProperty('color', $header_style) ? $header_style['color'] : '';
        $header_text_align  = Helper::hasArrayProperty('align', $header_style) ? $header_style['align'] : self::TEXT_ALIGN_HORIZONTAL_CENTER;

        $col_total  = count($header_titles) > 0 ? count($header_titles) : 0;


        $align      = Helper::hasArrayProperty('align', $params) ? $params['align'] : self::TEXT_ALIGN_HORIZONTAL_CENTER;
        $text       = Helper::hasArrayProperty('text', $params) ? $params['text'] : '';
        $bold       = Helper::hasArrayProperty('bold', $params) ? $params['bold'] : false;
        $color      = Helper::hasArrayProperty('color', $params) ? $params['color'] : self::DEFAULT_COLOR;
        $size       = Helper::hasArrayProperty('size', $params) ? $params['size'] : self::TEXT_SIZE;

        // get current slide
        $currentSlide = $this->_presentation->getActiveSlide();

        // Create a shape (table)
        $shape = $currentSlide->createTableShape($col_total);
        $shape->setHeight($height);
        $shape->setWidth($width);
        $shape->setOffsetX($offset_x);
        $shape->setOffsetY($offset_y);
        $shape->getBorder()->setColor(new Color('FFFFFFFF'))->setLineWidth(1);

        // Add row header
        $row = $shape->createRow();
        $row->setHeight(20);
        $row->getFill()->setFillType(Fill::FILL_SOLID)
            ->setRotation(90)
            ->setStartColor(new Color($header_background))
            ->setEndColor(new Color($header_background));

        foreach ($header_titles as $header_title) {
            $row->nextCell()->createTextRun($header_title)->getFont()->setBold($bold);
            $paragraph = $row->getCell()->getActiveParagraph();
            Alignment::setAlignText($paragraph, $header_text_align);
        }

        foreach ($row->getCells() as $cell) {
            $cell->getBorders()->getBottom()->setLineWidth(1)
                ->setColor(new Color('FFFFFFFF'))
                ->setLineStyle(Border::LINE_SINGLE)
                ->setDashStyle(Border::DASH_SOLID);
        }


        $row = $shape->createRow();
        $row->getFill()->setFillType(Fill::FILL_SOLID)
            ->setRotation(90)
            ->setStartColor(new Color('FFE06B20'))
            ->setEndColor(new Color('FFFFFFFF'));
        $cell = $row->nextCell();
        $cell->setColSpan(3);
        $cell->createTextRun('Title row')->getFont()->setBold(true)->setSize(16);
        $cell->getBorders()->getBottom()->setLineWidth(4)
            ->setLineStyle(Border::LINE_SINGLE)
            ->setDashStyle(Border::DASH_DASH);

// Add row
//        echo date('H:i:s') . ' Add row'.EOL;
        $row = $shape->createRow();
        $row->setHeight(20);
        $row->getFill()->setFillType(Fill::FILL_SOLID)
            ->setRotation(90)
            ->setStartColor(new Color('FFE06B20'))
            ->setEndColor(new Color('FFFFFFFF'));
        $row->nextCell()->createTextRun('R1C1')->getFont()->setBold(true);
        $row->nextCell()->createTextRun('R1C2')->getFont()->setBold(true);
        $row->nextCell()->createTextRun('R1C3')->getFont()->setBold(true);

        foreach ($row->getCells() as $cell) {
            $cell->getBorders()->getTop()->setLineWidth(4)
                ->setLineStyle(Border::LINE_SINGLE)
                ->setDashStyle(Border::DASH_DASH);
        }

// Add row
//        echo date('H:i:s') . ' Add row'.EOL;
        $row = $shape->createRow();
        $row->getFill()->setFillType(Fill::FILL_SOLID)
            ->setStartColor(new Color('FFE06B20'))
            ->setEndColor(new Color('FFE06B20'));
        $row->nextCell()->createTextRun('R2C1');
        $row->nextCell()->createTextRun('R2C2');
        $row->nextCell()->createTextRun('R2C3');

// Add row
//        echo date('H:i:s') . ' Add row'.EOL;
        $row = $shape->createRow();
        $row->getFill()->setFillType(Fill::FILL_SOLID)
            ->setStartColor(new Color('FFE06B20'))
            ->setEndColor(new Color('FFE06B20'));
        $row->nextCell()->createTextRun('R3C1');
        $row->nextCell()->createTextRun('R3C2');
        $row->nextCell()->createTextRun('R3C3');

// Add row
//        echo date('H:i:s') . ' Add row'.EOL;
        $row = $shape->createRow();
        $row->getFill()->setFillType(Fill::FILL_SOLID)
            ->setStartColor(new Color('FFE06B20'))
            ->setEndColor(new Color('FFE06B20'));
        $cellC1 = $row->nextCell();
        $textRunC1 = $cellC1->createTextRun('Link');
        $textRunC1->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/')->setTooltip('PHPPresentation');
        $cellC2 = $row->nextCell();
        $textRunC2 = $cellC2->createTextRun('RichText with');
        $textRunC2->getFont()->setBold(true);
        $textRunC2->getFont()->setSize(12);
        $textRunC2->getFont()->setColor(new Color('FF000000'));
        $cellC2->createBreak();
        $textRunC2 = $cellC2->createTextRun('Multiline');
        $textRunC2->getFont()->setBold(true);
        $textRunC2->getFont()->setSize(14);
        $textRunC2->getFont()->setColor(new Color('FF0088FF'));
        $cellC3 = $row->nextCell();
        $textRunC3 = $cellC3->createTextRun('Link Github');
        $textRunC3->getHyperlink()->setUrl('https://github.com')->setTooltip('GitHub');
        $cellC3->createBreak();
        $textRunC3 = $cellC3->createTextRun('Link Google');
        $textRunC3->getHyperlink()->setUrl('https://google.com')->setTooltip('Google');


    }


    //ASSESORS

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
            if(!empty($slide['texts'])){
                if(Helper::is_multi_array($slide['texts'])){
                    foreach ($slide['texts'] as $item) {
                        $this->createText($item);
                    }
                }else{
                    $this->createText($slide['texts']);
                }
            }

            //add image
            if(!empty($slide['images'])) {
                if (Helper::is_multi_array($slide['images'])) {
                    foreach ($slide['images'] as $item) {
                        $this->createImage($item);
                    }
                } else {
                    $this->createImage($slide['images']);
                }
            }

            //add table
            if(!empty($slide['tables'])) {
                if (Helper::is_multi_array($slide['tables'])) {
                    foreach ($slide['tables'] as $item) {
                        $this->createTable($item);
                    }
                } else {
                    $this->createTable($slide['tables']);
                }
            }
        }
    }
}