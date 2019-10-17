<?php
/**
 * Creado con PhpStorm.
 * Copyright (c) html 2016.
 * Autor: Franklyn Alejandro Sosa Pérez <alesjohnson@hotmail.com>
 * Fecha: 21/09/2016
 * Hora: 9:59
 */

namespace AlejandroSosa\YiiPowerPoint\Common;

use PhpOffice\PhpPresentation\Slide;
use AlejandroSosa\YiiPowerPoint\Common\Style;
use AlejandroSosa\YiiPowerPoint\Common\Helper;

/**
 * Class Texts
 * @package AlejandroSosa\YiiPowerPoint\Common
 */
class Texts extends AbstractObject 
{
    /**
     * Create custom object
     * @param Slide $slide
     * @param array $options
     * @return mixed
     */
    public static function create(Slide $slide, $options = array())
    {
        //check if options is only one or multiple
        if (Helper::isArrayMultidimensional($options)) {
            foreach ($options as $item) {
                self::createCustomText($slide, $item);
            }
        } else {
            self::createCustomText($slide, $options);
        }
    }

    /**
     * Create object text into slide
     * @param Slide $slide
     * @param array $params
     */
    private static function createCustomText(Slide $slide, $params = array())
    {
        $text       = Helper::hasArrayProperty('text', $params) ? $params['text'] : '';
        $options    = Helper::hasArrayProperty('options', $params) ? $params['options'] : array();

        $height     = Helper::hasArrayProperty('height', $options) ? $options['height'] : self::TEXT_HEIGHT;
        $width      = Helper::hasArrayProperty('width', $options) ? $options['width'] : self::TEXT_WIDTH;
        $offset_x   = Helper::hasArrayProperty('ox', $options) ? $options['ox'] : self::TEXT_OFFSET_X;
        $offset_y   = Helper::hasArrayProperty('oy', $options) ? $options['oy'] : self::TEXT_OFFSET_Y;
        $align      = Helper::hasArrayProperty('align', $options) ? $options['align'] : self::TEXT_ALIGN_HORIZONTAL_CENTER;
        $bold       = Helper::hasArrayProperty('bold', $options) ? $options['bold'] : false;
        $color      = Helper::hasArrayProperty('color', $options) ? $options['color'] : self::COLOR_PRIMARY_TEXT;
        $size       = Helper::hasArrayProperty('size', $options) ? $options['size'] : self::TEXT_SIZE;

        $current_slide = $slide;
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
}
