<?php
/**
 * Creado con PhpStorm.
 * Copyright (c) html 2016.
 * Autor: Franklyn Alejandro Sosa Pérez <alesjohnson@hotmail.com>
 * Fecha: 14/09/2016
 * Hora: 16:05
 */

namespace AlejandroSosa\YiiPowerPoint\Common;

use AlejandroSosa\YiiPowerPoint\PowerPoint;
use PhpOffice\PhpPresentation\Style\Alignment as Align;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;

/**
 * Class Alignment
 * @package AlejandroSosa\YiiPowerPoint\Common
 */
class Alignment extends PowerPoint
{
    /**
     * Set align of text
     * @param Paragraph $shape
     * @param $align
     */
    public static function setAlignText(Paragraph $shape, $align)
    {
        if(!empty($shape)){
            $obj = $shape->getAlignment();

            switch ($align){
                //horizontal
                case self::TEXT_ALIGN_HORIZONTAL_GENERAL: $obj->setHorizontal(Align::HORIZONTAL_GENERAL); break;
                case self::TEXT_ALIGN_HORIZONTAL_CENTER: $obj->setHorizontal(Align::HORIZONTAL_CENTER); break;
                case self::TEXT_ALIGN_HORIZONTAL_LEFT: $obj->setHorizontal(Align::HORIZONTAL_LEFT); break;
                case self::TEXT_ALIGN_HORIZONTAL_RIGHT: $obj->setHorizontal(Align::HORIZONTAL_RIGHT); break;
                case self::TEXT_ALIGN_HORIZONTAL_JUSTIFY: $obj->setHorizontal(Align::HORIZONTAL_JUSTIFY); break;
                case self::TEXT_ALIGN_HORIZONTAL_DISTRIBUTED: $obj->setHorizontal(Align::HORIZONTAL_DISTRIBUTED); break;
                //vertical
                case self::TEXT_ALIGN_VERTICAL_AUTO: $obj->setVertical(Align::VERTICAL_AUTO); break;
                case self::TEXT_ALIGN_VERTICAL_CENTER: $obj->setVertical(Align::VERTICAL_CENTER); break;
                case self::TEXT_ALIGN_VERTICAL_TOP: $obj->setVertical(Align::VERTICAL_TOP); break;
                case self::TEXT_ALIGN_VERTICAL_BOTTOM: $obj->setVertical(Align::VERTICAL_BOTTOM); break;
                case self::TEXT_ALIGN_VERTICAL_BASE: $obj->setVertical(Align::VERTICAL_BASE); break;

                default: $obj->setHorizontal(Align::HORIZONTAL_CENTER);
            }
        }
    }
}