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

/**
 * Class Texts
 * @package AlejandroSosa\YiiPowerPoint\Common
 */
class Texts extends AbstractObject
{
    /**
     * @param Slide $slide
     * @param array $options
     * @return mixed
     */
    public static function create(Slide $slide, $options = [])
    {
        // TODO: Implement create() method.
        return 'hola desde texts';
    }
}