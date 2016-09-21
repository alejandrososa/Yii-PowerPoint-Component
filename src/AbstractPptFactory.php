<?php
/**
 * Creado con PhpStorm.
 * Copyright (c) html 2016.
 * Autor: Franklyn Alejandro Sosa Pérez <alesjohnson@hotmail.com>
 * Fecha: 21/09/2016
 * Hora: 10:30
 */

namespace AlejandroSosa\YiiPowerPoint;
use PhpOffice\PhpPresentation\Slide;

/**
 * Class AbstractPptFactory
 * @package AlejandroSosa\YiiPowerPoint
 */
abstract class AbstractPptFactory
{
    /**
     * Create objects dynamically to add a slide
     * @param Slide $slide
     * @param array $options
     * @return mixed
     */
    abstract public static function build(Slide $slide, $options = []);
}