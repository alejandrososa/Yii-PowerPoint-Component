<?php

/**
 * Creado con PhpStorm.
 * PowerPoint
 * Autor: Franklyn Alejandro Sosa Pérez <alesjohnson@hotmail.com>
 * Fecha: 8/9/2016
 * Hora: 22:13
 */

namespace AlejandroSosa\YiiPowerPoint\Common;

/**
 * Class Helper
 * @package AlejandroSosa\YiiPowerPoint\Common
 */
class Helper
{
    /**
     * Check if array is multidimensional
     * @param $arr
     * @return bool
     */
    public static function isMultiArray( $arr ) {
        rsort( $arr );
        return isset( $arr[0] ) && is_array( $arr[0] );
    }

    /**
     * Check if array is multidimensional
     * @param $arr
     * @return bool
     */
    public static function isArrayMultidimensional( $arr ) {
        if (!is_array($arr)) {
            return false;
        }
        foreach ($arr as $elm) {
            if (!is_array($elm)){
                return false;
            }
        }
        return true;
    }

    /**
     * Check if attribute options has property
     * @param string $property
     * @param array $array
     * @return bool
     */
    public static function hasArrayProperty($property, $array = [])
    {
        return !empty($array) && array_key_exists($property, $array) ? true : false;
    }

    /**
     * Convert string to array by delimiter
     * @param $string
     * @param $delimiter
     * @return array
     */
    public static function convertStringToArray($string, $delimiter)
    {
        if(empty($string) && empty($delimiter)){
            return [];
        }
        return explode($delimiter, $string);
    }

    /**
     * Check if string contains substring
     * @param $string
     * @param $substring
     * @return bool
     */
    public static function stringContains($string, $substring)
    {
        if(empty($string) && empty($substring)){
            return false;
        }
        return strpos($string, $substring) !== false ? true : false;
    }

    /**
     * Create directory if not exists
     * @param $path_directory
     */
    public static function createDirectory($path_directory)
    {
        if(isset($path_directory)) {
            if (!file_exists($path_directory)) {
                mkdir($path_directory);
            }
        }
    }
}