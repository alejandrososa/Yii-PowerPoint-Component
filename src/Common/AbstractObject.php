<?php
/**
 * Creado con PhpStorm.
 * Copyright (c) html 2016.
 * Autor: Franklyn Alejandro Sosa Pérez <alesjohnson@hotmail.com>
 * Fecha: 21/09/2016
 * Hora: 10:43
 */

namespace AlejandroSosa\YiiPowerPoint\Common;
use PhpOffice\PhpPresentation\Slide;

/**
 * Class AbstractObject
 * @package AlejandroSosa\YiiPowerPoint
 */
abstract class AbstractObject implements ConstantesPPT
{
    /**
     * Create custom object
     * @param Slide $slide
     * @param array $options
     * @return mixed
     */
    abstract public static function create(Slide $slide, $options = array());
}