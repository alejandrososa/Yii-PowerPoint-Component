<?php
/**
 * Creado con PhpStorm.
 * Copyright (c) html 2016.
 * Autor: Franklyn Alejandro Sosa Pérez <alesjohnson@hotmail.com>
 * Fecha: 20/09/2016
 * Hora: 12:14
 */

namespace AlejandroSosa\YiiPowerPoint\Common;

use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Shape\RichText\Run;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Shape\Table\Cell;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Slide\Background\Image;

/**
 * Class Style
 * @package AlejandroSosa\YiiPowerPoint\Common
 */
class Style
{
    /**
     * Set background of slide
     * @param Slide $slide
     * @param $path_image
     */
    public static function setBackgroundSlide(Slide $slide, $path_image)
    {
        if($slide instanceof Slide && !empty($path_image)) {
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
    public static function setStyleText(Run $text, $size = 10, $bold = false, $color = 'FF000000')
    {
        if($text instanceof Run) {
            $text->getFont()->setBold($bold);
            $text->getFont()->setSize($size);
            $text->getFont()->setColor(new Color($color));
        }
    }

    /**
     * Set background of column
     * @param Cell $column
     * @param string $color
     */
    public static function setBackgroundColumn(Cell $column, $color = 'FF000000')
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
    public static function setBorderColumn(Cell $column, $color = 'FFFFFFFF', $width = 1)
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
}