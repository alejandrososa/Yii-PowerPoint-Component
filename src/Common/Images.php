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
use PhpOffice\PhpPresentation\Shape\Drawing;
/**
 * Class Images
 * @package AlejandroSosa\YiiPowerPoint\Common
 */
class Images extends AbstractObject
{
    /**
     * Create custom object
     * @param Slide $slide
     * @param array $options
     * @return mixed
     */
    public static function create(Slide $slide, $options = [])
    {
        // TODO: Implement create() method.
    }

    /**
     * Create custom image into slide
     * @param array $params
     */
    private function createCustomImage(Slide $slide, $params = [])
    {
        $height     = Helper::hasArrayProperty('height', $params) ? $params['height'] : self::TEXT_HEIGHT;
        $width      = Helper::hasArrayProperty('width', $params) ? $params['width'] : self::TEXT_WIDTH;
        $offset_x   = Helper::hasArrayProperty('ox', $params) ? $params['ox'] : self::TEXT_OFFSET_X;
        $offset_y   = Helper::hasArrayProperty('oy', $params) ? $params['oy'] : self::TEXT_OFFSET_Y;
        $name       = Helper::hasArrayProperty('name', $params) ? $params['name'] : '';
        $description= Helper::hasArrayProperty('decription', $params) ? $params['name'] : '';

//        if()

        $current_slide = $slide;
//        $shape = $current_slide->createRichTextShape();

        $shape = new Drawing\File();
        $shape->setName($name)->setDescription($description);
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