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
use AlejandroSosa\YiiPowerPoint\AbstractPptFactory;
use AlejandroSosa\YiiPowerPoint\Common\Charts;
use AlejandroSosa\YiiPowerPoint\Common\Images;
use AlejandroSosa\YiiPowerPoint\Common\Tables;
use AlejandroSosa\YiiPowerPoint\Common\Texts;

/**
 * Class ObjectsPptFactory
 * @package AlejandroSosa\YiiPowerPoint
 */
class ObjectsPptFactory extends AbstractPptFactory
{
    /**
     * Create objects dynamically to add a slide
     * @param $type type of object (text, image, table, chart)
     * @param Slide $slide
     * @param array $options
     * @return mixed
     */
    public static function build($type, Slide $slide, $options = [])
    {
        $obj = NULL;
        switch ($type) {
            case self::TYPE_TEXT:
                $obj = Texts::create($slide, $options);
                break;
            case self::TYPE_IMAGE:
                $obj = Images::create($slide, $options);
                break;
            case self::TYPE_TABLE:
                $obj = Tables::create($slide, $options);
                break;
            case self::TYPE_CHART:
                $obj = Charts::create($slide, $options);
                break;
        }
        return $obj;
    }
}