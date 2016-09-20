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
use AlejandroSosa\YiiPowerPoint\Common\Alignment;
use AlejandroSosa\YiiPowerPoint\Common\Style;

/**
 * Class Table
 * @package AlejandroSosa\YiiPowerPoint\Common
 */
class Table
{
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
    public static function createTable(Slide $slide, $col_total, $height, $width, $offset_x, $offset_y)
    {
        $shape = $slide->createTableShape($col_total);
        $shape->setHeight($height);
        $shape->setWidth($width);
        $shape->setOffsetX($offset_x);
        $shape->setOffsetY($offset_y);
        return $shape;
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
    public static function createRow(TBL $table, $texts = [], $size = 10, $bold = false, $color = 'FF000000',
                                     $align, $background = 'FFFFFFFF', $width = 100, $height = 20){

        $align = !empty($align) ? $align : Alignment::TEXT_ALIGN_HORIZONTAL_CENTER;

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
            if(is_array($text) && Helper::isMultiArray($text)){
                $style  = Helper::hasArrayProperty('style', $text) ? $text['style'] : [];
                $_text  = Helper::hasArrayProperty('text', $text) ? $text['text'] : '';
                $_bold  = Helper::hasArrayProperty('bold', $style) ? $style['bold'] : $bold;
                $_size  = Helper::hasArrayProperty('size', $style) ? $style['size'] : $size;
                $_color = Helper::hasArrayProperty('color', $style) ? $style['color'] : $color;
                $_align = Helper::hasArrayProperty('align', $style) ? $style['align'] : $align;
                $_width = Helper::hasArrayProperty('width', $style) ? $style['width'] : $width;
                $_bkg   = Helper::hasArrayProperty('background', $style) ? $style['background'] : $background;

                //set style text
                $current_text = $current_column->createTextRun($_text);
                Style::setStyleText($current_text, $_size, $_bold,$_color);

                //set align text
                $paragraph = $current_column->getActiveParagraph();
                Alignment::setAlignText($paragraph, $_align);

                //set background
                Style::setBackgroundColumn($current_column, $_bkg);

                //set width
                Style::setWidthColumn($_width);
            }else{
                //whe text not have options
                //set style text
                $current_text = $current_column->createTextRun($text);
                Style::setStyleText($current_text, $size, $bold, $color);

                //set align text
                $paragraph = $current_column->getActiveParagraph();
                Alignment::setAlignText($paragraph, $align);
            }
        }
    }
}