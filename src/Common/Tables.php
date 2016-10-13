<?php
/**
 * Creado con PhpStorm.
 * Copyright (c) html 2016.
 * Autor: Franklyn Alejandro Sosa Pérez <alesjohnson@hotmail.com>
 * Fecha: 16/09/2016
 * Hora: 14:46
 */

namespace AlejandroSosa\YiiPowerPoint\Common;

use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Shape\Table as TBL;
use PhpOffice\PhpPresentation\Shape\Table\Row;
use PhpOffice\PhpPresentation\Shape\Table\Cell;
use PhpOffice\PhpPresentation\Shape\RichText\Run;
use AlejandroSosa\YiiPowerPoint\Common\Style;

/**
 * Class Table
 * @package AlejandroSosa\YiiPowerPoint\Common
 */
class Tables extends AbstractObject
{
    /**
     * Create custom object
     * @param Slide $slide
     * @param array $options
     * @return mixed
     */
    public static function create(Slide $slide, $options = [])
    {
        //check if options is only one or multiple
        if (Helper::isArrayMultidimensional($options)) {
            foreach ($options as $item) {
                self::createCustomTable($slide, $item);
            }
        } else {
            self::createCustomTable($slide, $options);
        }
    }

    /**
     * Create custom table into slide
     * @param Slide $slide
     * @param array $params
     */
    private function createCustomTable(Slide $slide, $params = [])
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
        $header_background  = Helper::hasArrayProperty('background', $header_style) ? $header_style['background'] : self::COLOR_WHITE;
        $header_text_bold   = Helper::hasArrayProperty('bold', $header_style) ? $header_style['bold'] : false;
        $header_text_size   = Helper::hasArrayProperty('size', $header_style) ? $header_style['size'] : self::TEXT_SIZE;
        $header_text_color  = Helper::hasArrayProperty('color', $header_style) ? $header_style['color'] : self::DEFAULT_COLOR;
        $header_text_align  = Helper::hasArrayProperty('align', $header_style) ? $header_style['align'] : self::TEXT_ALIGN_HORIZONTAL_CENTER;
        $header_width       = Helper::hasArrayProperty('width', $header_style) ? $header_style['width'] : 100;
        $header_height      = Helper::hasArrayProperty('height', $header_style) ? $header_style['height'] : 20;

        $col_total  = count($header_columns) > 0 ? count($header_columns) : 0;

        //get current slide
        $current_slide = $slide;

        //create a table shape
        $shape = self::makeTable($current_slide, $col_total, $height, $width, $offset_x, $offset_y);

        //add row header
        self::makeRow($shape, $header_columns, $header_text_size, $header_text_bold, $header_text_color,
            $header_text_align, $header_background, $header_width, $header_height);

        //add the remaining rows
        foreach ($rows as $row) {
            $texts  = !empty($row['columns']) ? $row['columns'] : [];
            $style  = !empty($row['style']) ? $row['style'] : [];
            $size   = !empty($style['size']) ? $style['size'] : self::TEXT_SIZE;
            $bold   = !empty($style['bold']) ? $style['bold'] : self::FALSE;
            $color  = !empty($style['color']) ? $style['color'] : self::COLOR_PRIMARY_TEXT;
            $align  = !empty($style['align']) ? $style['align'] : self::TEXT_ALIGN_HORIZONTAL_LEFT;
            $bg     = !empty($style['background']) ? $style['background'] : self::COLOR_WHITE;
            self::makeRow($shape, $texts, $size, $bold, $color, $align, $bg);
        }
    }

    /**
     * Create table
     * @param Slide $slide
     * @param $col_total
     * @param $height
     * @param $width
     * @param $offset_x
     * @param $offset_y
     * @return TBL
     */
    private function makeTable(Slide $slide, $col_total, $height, $width, $offset_x, $offset_y)
    {
        if($slide instanceof Slide) {
            $shape = $slide->createTableShape($col_total);
            $shape->setHeight($height);
            $shape->setWidth($width);
            $shape->setOffsetX($offset_x);
            $shape->setOffsetY($offset_y);
            return $shape;
        }
    }

    /**
     * Create custom row
     * @param TBL $table
     * @param array $texts
     * @param int $size
     * @param bool $bold
     * @param string $color
     * @param $align
     * @param string $background
     * @param int $width
     * @param int $height
     */
    private function makeRow(TBL $table, $texts = [], $size = 10, $bold = false, $color = 'FF000000',
                             $align, $background = 'FFFFFFFF', $width = 100, $height = 20){

        $align = !empty($align) ? $align : self::TEXT_ALIGN_HORIZONTAL_CENTER;

        if($table instanceof TBL) {
            //create row
            $row = $table->createRow();
            $row->setHeight($height);
            $row->getFill()->setFillType(Fill::FILL_SOLID)
                ->setStartColor(new Color($background))
                ->setEndColor(new Color($background));


            //set each columns
            foreach ($texts as $text) {
                $current_column = $row->nextCell();

                //set borders
                Style::setBorderColumn($current_column);

                //when text has options
                if (is_array($text) && Helper::isMultiArray($text)) {
                    $style = Helper::hasArrayProperty('style', $text) ? $text['style'] : [];
                    $_text = Helper::hasArrayProperty('text', $text) ? $text['text'] : '';
                    $_bold = Helper::hasArrayProperty('bold', $style) ? $style['bold'] : $bold;
                    $_size = Helper::hasArrayProperty('size', $style) ? $style['size'] : $size;
                    $_color = Helper::hasArrayProperty('color', $style) ? $style['color'] : $color;
                    $_align = Helper::hasArrayProperty('align', $style) ? $style['align'] : $align;
                    $_width = Helper::hasArrayProperty('width', $style) ? $style['width'] : $width;
                    $_bkg = Helper::hasArrayProperty('background', $style) ? $style['background'] : $background;

                    //set style text
                    $current_text = $current_column->createTextRun($_text);
                    Style::setStyleText($current_text, $_size, $_bold, $_color);

                    //set align text
                    $paragraph = $current_column->getActiveParagraph();
                    Style::setAlignText($paragraph, $_align);

                    //set background
                    Style::setBackgroundColumn($current_column, $_bkg);

                    //set width
                    Style::setWidthColumn($current_column, $_width);
                } else {
                    //whe text not have options
                    //set style text
                    $current_text = $current_column->createTextRun($text);
                    Style::setStyleText($current_text, $size, $bold, $color);

                    //set align text
                    $paragraph = $current_column->getActiveParagraph();
                    Style::setAlignText($paragraph, $align);
                }
            }
        }
    }
}