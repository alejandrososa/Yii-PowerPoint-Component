<?php
/**
 * Creado con PhpStorm.
 * Copyright (c) html 2016.
 * Autor: Franklyn Alejandro Sosa Pérez <alesjohnson@hotmail.com>
 * Fecha: 21/09/2016
 * Hora: 10:21
 */

namespace AlejandroSosa\YiiPowerPoint;

use AlejandroSosa\YiiPowerPoint\Common\Charts;
use AlejandroSosa\YiiPowerPoint\Common\Images;
use AlejandroSosa\YiiPowerPoint\Common\Tables;
use AlejandroSosa\YiiPowerPoint\Common\Texts;

/**
 * Class ObjectFactory
 * @package AlejandroSosa\YiiPowerPoint
 */
class ObjectFactory extends AbstractFactory
{
    /**
     * Build object
     * @param array $slide
     * @return Charts|Images|Tables|Texts|null
     */
    public function build($slide = [])
    {
        $object = NULL;
        switch ($slide) {
            case "texts":
                $object = new Texts();
                break;
            case "images":
                $object = new Images();
                break;
            case "tables":
                $object = new Tables();
                break;
            case "charts":
                $object = new Charts();
                break;
        }
        return $object;
    }
}