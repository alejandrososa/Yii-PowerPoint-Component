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
    const TYPE_TEXT     = 'texts';
    const TYPE_IMAGE    = 'images';
    const TYPE_TABLE    = 'tables';
    const TYPE_CHART    = 'charts';

    /**
     * Create objects dynamically to add a slide
     * @param $type type of object (text, image, table, chart)
     * @param Slide $slide
     * @param array $options
     * @return mixed
     */
    abstract public static function build($type, Slide $slide, $options = array());
}