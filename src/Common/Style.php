<?php
/**
 * Creado con PhpStorm.
 * Copyright (c) html 2016.
 * Autor: Franklyn Alejandro Sosa Pérez <alesjohnson@hotmail.com>
 * Fecha: 20/09/2016
 * Hora: 12:14
 */

namespace AlejandroSosa\YiiPowerPoint\Common;

use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\Background\Image;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\Run;
use PhpOffice\PhpPresentation\Shape\RichText\Paragraph;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Shape\Table\Cell;

use AlejandroSosa\YiiPowerPoint\PowerPoint;


/**
 * Class Style
 * @package AlejandroSosa\YiiPowerPoint\Common
 */
class Style implements ConstantesPPT
{
    /**
     * Set background of slide
     * @param Slide $slide
     * @param $path_image
     */
    public static function setBackgroundSlide(Slide $slide, $path_image)
    {
        if($slide instanceof Slide && file_exists($path_image)) {
            $image = new Image();
            $image->setPath($path_image);
            $slide->setBackground($image);
        }
    }

    /**
     * Set style of text run
     * @param Run $text
     * @param int $size
     * @param bool $bold
     * @param string $color
     */
    public static function setStyleText(Run $text, $size = 10, $bold = false, $color = self::COLOR_PRIMARY_TEXT)
    {
        if($text instanceof Run) {
            $text->getFont()->setBold($bold);
            $text->getFont()->setSize($size);
            $text->getFont()->setColor(new Color($color));
        }
    }

    /**
     * Set align of text
     * @param Paragraph $shape
     * @param $align
     */
    public static function setAlignText(Paragraph $shape, $align)
    {
        if($shape instanceof Paragraph) {
            $obj = $shape->getAlignment();

            switch ($align){
                //horizontal
                case self::TEXT_ALIGN_HORIZONTAL_GENERAL: $obj->setHorizontal(Alignment::HORIZONTAL_GENERAL); break;
                case self::TEXT_ALIGN_HORIZONTAL_CENTER: $obj->setHorizontal(Alignment::HORIZONTAL_CENTER); break;
                case self::TEXT_ALIGN_HORIZONTAL_LEFT: $obj->setHorizontal(Alignment::HORIZONTAL_LEFT); break;
                case self::TEXT_ALIGN_HORIZONTAL_RIGHT: $obj->setHorizontal(Alignment::HORIZONTAL_RIGHT); break;
                case self::TEXT_ALIGN_HORIZONTAL_JUSTIFY: $obj->setHorizontal(Alignment::HORIZONTAL_JUSTIFY); break;
                case self::TEXT_ALIGN_HORIZONTAL_DISTRIBUTED: $obj->setHorizontal(Alignment::HORIZONTAL_DISTRIBUTED); break;
                //vertical
                case self::TEXT_ALIGN_VERTICAL_AUTO: $obj->setVertical(Alignment::VERTICAL_AUTO); break;
                case self::TEXT_ALIGN_VERTICAL_CENTER: $obj->setVertical(Alignment::VERTICAL_CENTER); break;
                case self::TEXT_ALIGN_VERTICAL_TOP: $obj->setVertical(Alignment::VERTICAL_TOP); break;
                case self::TEXT_ALIGN_VERTICAL_BOTTOM: $obj->setVertical(Alignment::VERTICAL_BOTTOM); break;
                case self::TEXT_ALIGN_VERTICAL_BASE: $obj->setVertical(Alignment::VERTICAL_BASE); break;

                default: $obj->setHorizontal(Alignment::HORIZONTAL_CENTER);
            }

            //set align margin
            $obj->setMarginLeft(0);
            $obj->setMarginRight(0);
        }
    }

    /**
     * Set background of column
     * @param Cell $column
     * @param string $color
     */
    public static function setBackgroundColumn(Cell $column, $color = self::COLOR_WHITE)
    {
        if($column instanceof Cell) {
            $column->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->setStartColor(new Color($color))
                ->setEndColor(new Color($color));
        }
    }

    /**
     * Set border of column
     * @param Cell $column
     * @param string $color
     * @param int $width
     */
    public static function setBorderColumn(Cell $column, $color = self::COLOR_GREY, $width = 1)
    {
        if($column instanceof Cell){
            $column->getBorders()->getBottom()->setColor(new Color($color))->setLineWidth($width);
            $column->getBorders()->getTop()->setColor(new Color($color))->setLineWidth($width);
            $column->getBorders()->getLeft()->setColor(new Color($color))->setLineWidth($width);
            $column->getBorders()->getRight()->setColor(new Color($color))->setLineWidth($width);
        }
    }

    /**
     * Set width of column
     * @param Cell $column
     * @param int $width
     */
    public static function setWidthColumn(Cell $column, $width = 100)
    {
        if($column instanceof Cell){
            $column->setWidth($width);
        }
    }

    /**
     * Get rgb color
     * @param $color
     * @return mixed
     */
    public static function getRGB($color){
        return strlen($color) > 6 ? substr($color, 2) : $color;
    }
}