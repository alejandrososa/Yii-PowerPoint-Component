<?php
/**
 * Creado con PhpStorm.
 * Copyright (c) html 2016.
 * Autor: Franklyn Alejandro Sosa Pérez <alesjohnson@hotmail.com>
 * Fecha: 14/09/2016
 * Hora: 16:23
 */

namespace AlejandroSosa\YiiPowerPoint\Common;
use PhpOffice\PhpPresentation\Slide;

/**
 * Class Images
 * @package AlejandroSosa\YiiPowerPoint\Common
 */
class Images extends AbstractObject
{
    /**
     * @param Slide $slide
     * @param array $options
     * @return mixed
     */
    public static function create(Slide $slide, $options = [])
    {
        // TODO: Implement create() method.
        return 'hola desde images';
    }

    /**
     * Check if image base64
     * @param $base64 string
     * @return bool
     */
    public function check_base64_image($base64) {
        $result = false;
        $img = imagecreatefromstring(base64_decode($base64));
        if (!$img) {
            return $result;
        }

        imagepng($img, 'tmp.png');
        $info = getimagesize('tmp.png');

        unlink('tmp.png');

        if ($info[0] > 0 && $info[1] > 0 && $info['mime']) {
            $result = $info;
        }

        return $result;
    }

}