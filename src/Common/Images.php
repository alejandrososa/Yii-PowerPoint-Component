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
    public static function create(Slide $slide, $options = array())
    {
        //check if options is only one or multiple
        if (Helper::isArrayMultidimensional($options)) {
            foreach ($options as $item) {
                self::createCustomImage($slide, $item);
            }
        } else {
            self::createCustomImage($slide, $options);
        }
    }

    /**
     * Create custom image into slide
     * @param array $params
     */
    private function createCustomImage(Slide $slide, $params = array())
    {
        $height         = Helper::hasArrayProperty('height', $params) ? $params['height'] : self::TEXT_HEIGHT;
        $width          = Helper::hasArrayProperty('width', $params) ? $params['width'] : self::TEXT_WIDTH;
        $offset_x       = Helper::hasArrayProperty('ox', $params) ? $params['ox'] : self::TEXT_OFFSET_X;
        $offset_y       = Helper::hasArrayProperty('oy', $params) ? $params['oy'] : self::TEXT_OFFSET_Y;
        $name           = Helper::hasArrayProperty('name', $params) ? $params['name'] : '';
        $description    = Helper::hasArrayProperty('decription', $params) ? $params['decription'] : '';
        $path           = Helper::hasArrayProperty('path', $params) ? $params['path'] : '';

        $shape = new Drawing\File();
        $shape->setName($name)->setDescription($description);

        $shape = new Drawing\File();
        $shape->setName($name);
        $shape->setDescription($description);
        $shape->setPath($path);

        if(!empty($width) && !empty($height)){
            $shape->setWidthAndHeight($width, $height);
        }else{
            $shape->setHeight($height);
        }

        $shape->setOffsetX($offset_x);
        $shape->setOffsetY($offset_y);
        $slide->addShape($shape);
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