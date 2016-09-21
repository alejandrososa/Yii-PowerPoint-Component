<?php
/**
 * Creado con PhpStorm.
 * Copyright (c) html 2016.
 * Autor: Franklyn Alejandro Sosa Pérez <alesjohnson@hotmail.com>
 * Fecha: 21/09/2016
 * Hora: 10:30
 */

namespace AlejandroSosa\YiiPowerPoint;

/**
 * Class AbstractFactory
 * @package AlejandroSosa\YiiPowerPoint
 */
abstract class AbstractFactory
{
    abstract function build($slide = []);
}