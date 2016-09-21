<?php
/**
 * Creado con PhpStorm.
 * Copyright (c) html 2016.
 * Autor: Franklyn Alejandro Sosa Pérez <alesjohnson@hotmail.com>
 * Fecha: 21/09/2016
 * Hora: 10:21
 */

namespace AlejandroSosa\YiiPowerPoint;

use PhpOffice\PhpPresentation\Slide;
use AlejandroSosa\YiiPowerPoint\Common\Charts;
use AlejandroSosa\YiiPowerPoint\Common\Images;
use AlejandroSosa\YiiPowerPoint\Common\Tables;
use AlejandroSosa\YiiPowerPoint\Common\Texts;

/**
 * Class ObjectsPptFactory
 * @package AlejandroSosa\YiiPowerPoint
 */
class ObjectsPptFactory extends AbstractFactory
{
    /**
     * Create objects dynamically to add a slide
     * @param Slide $slide
     * @param array $options
     * @return mixed
     */
    public static function build(Slide $slide, $options = [])
    {
        $obj = NULL;
        switch ($slide) {
            case "texts":
                $obj = Texts::create($slide, $options);
                break;
            case "images":
                $obj = Images::create($slide, $options);
                break;
            case "tables":
                $obj = Tables::create($slide, $options);
                break;
            case "charts":
                $obj = Charts::create($slide, $options);
                break;
        }
        return $obj;
    }
}